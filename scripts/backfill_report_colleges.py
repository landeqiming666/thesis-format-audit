from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from app import (
    REPORTS_TABLE,
    UNKNOWN_COLLEGE,
    download_original_from_storage,
    get_supabase,
    safe_extract_college_from_docx,
)


def should_backfill(report: dict, include_existing: bool) -> bool:
    if report.get("status") != "success":
        return False
    if not (report.get("original_storage_path") or report.get("original_gcs_path")):
        return False
    if include_existing:
        return True
    return not report.get("college_name") or report.get("college_name") == UNKNOWN_COLLEGE


def fetch_reports() -> list[dict]:
    result = (
        get_supabase()
        .table(REPORTS_TABLE)
        .select(
            "id,status,original_filename,original_storage_path,original_gcs_path,"
            "college_name,college_source,college_raw_text,created_at"
        )
        .order("created_at", desc=False)
        .execute()
    )
    return result.data or []


def update_report_college(report_id: str, college_info: dict, dry_run: bool) -> None:
    if dry_run:
        return
    (
        get_supabase()
        .table(REPORTS_TABLE)
        .update(
            {
                "college_name": college_info["college_name"],
                "college_source": college_info["college_source"],
                "college_raw_text": college_info["college_raw_text"],
            }
        )
        .eq("id", report_id)
        .execute()
    )


def backfill_colleges(limit: int | None, include_existing: bool, dry_run: bool) -> dict:
    reports = [item for item in fetch_reports() if should_backfill(item, include_existing)]
    if limit is not None:
        reports = reports[:limit]

    stats = {"selected": len(reports), "updated": 0, "unknown": 0, "failed": 0, "skipped": 0}
    with tempfile.TemporaryDirectory(prefix="college-backfill-") as tmp:
        tmp_path = Path(tmp)
        for index, report in enumerate(reports, start=1):
            report_id = report["id"]
            original_name = report.get("original_filename") or "thesis.docx"
            try:
                original_bytes = download_original_from_storage(report)
                docx_path = tmp_path / f"{report_id}.docx"
                docx_path.write_bytes(original_bytes)
                college_info = safe_extract_college_from_docx(docx_path)
                update_report_college(report_id, college_info, dry_run)
                stats["updated"] += 1
                if college_info["college_name"] == UNKNOWN_COLLEGE:
                    stats["unknown"] += 1
                print(
                    f"[{index}/{len(reports)}] {report_id} {original_name} -> "
                    f"{college_info['college_name']} ({college_info['college_source'] or 'no source'})"
                )
            except Exception as exc:
                stats["failed"] += 1
                print(f"[{index}/{len(reports)}] {report_id} {original_name} -> FAILED: {exc}")
    return stats


def main() -> None:
    parser = argparse.ArgumentParser(description="Backfill per-report college classification from stored DOCX originals.")
    parser.add_argument("--limit", type=int, default=None, help="Only process the first N matching reports.")
    parser.add_argument("--include-existing", action="store_true", help="Recompute records that already have a college.")
    parser.add_argument("--dry-run", action="store_true", help="Download and classify, but do not update Supabase.")
    args = parser.parse_args()
    stats = backfill_colleges(args.limit, args.include_existing, args.dry_run)
    print(
        "Backfill finished: "
        f"selected={stats['selected']} updated={stats['updated']} "
        f"unknown={stats['unknown']} failed={stats['failed']} dry_run={args.dry_run}"
    )


if __name__ == "__main__":
    main()

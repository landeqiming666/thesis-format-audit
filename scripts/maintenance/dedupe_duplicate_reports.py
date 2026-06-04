from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import app


COUNTED_STATUSES = {"success", "storage_failed"}


def report_score(report: dict) -> tuple:
    """Prefer usable report records, then records with original files, then newest."""
    status_rank = {"success": 3, "storage_failed": 2, "audit_failed": 1}.get(report.get("status"), 0)
    has_report = 1 if report.get("report_storage_path") or report.get("report_gcs_path") else 0
    has_original = 1 if (
        report.get("original_storage_path")
        or report.get("original_gcs_path")
        or report.get("original_drive_file_id")
    ) else 0
    return (status_rank, has_report, has_original, report.get("created_at") or "")


def duplicate_groups(reports: list[dict]) -> list[dict]:
    groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for report in reports:
        sha = (report.get("original_sha256") or "").strip()
        if sha:
            groups[(report.get("user_id") or "", sha)].append(report)

    result = []
    for (_user_id, sha), rows in groups.items():
        if len(rows) <= 1:
            continue
        keep = max(rows, key=report_score)
        duplicates = sorted(
            [row for row in rows if row["id"] != keep["id"]],
            key=lambda row: row.get("created_at") or "",
            reverse=True,
        )
        result.append(
            {
                "sha256": sha,
                "user_id": keep.get("user_id") or "",
                "user_email": keep.get("user_email") or "",
                "original_filename": keep.get("original_filename") or "",
                "keep": keep,
                "duplicates": duplicates,
            }
        )
    return sorted(result, key=lambda item: len(item["duplicates"]), reverse=True)


def storage_paths_for_delete(report: dict) -> list[str]:
    paths = []
    for key in ("original_storage_path", "report_storage_path"):
        value = (report.get(key) or "").strip()
        if value:
            paths.append(value)
    return paths


def delete_supabase_storage(paths: list[str], execute: bool) -> dict:
    if not paths:
        return {"selected": 0, "deleted": 0, "failed": []}
    unique_paths = sorted(set(paths))
    if not execute:
        return {"selected": len(unique_paths), "deleted": 0, "failed": []}

    deleted = 0
    failed = []
    bucket = app.get_supabase().storage.from_(app.REPORTS_BUCKET)
    for index in range(0, len(unique_paths), 100):
        batch = unique_paths[index : index + 100]
        try:
            bucket.remove(batch)
            deleted += len(batch)
        except Exception as exc:
            failed.append({"paths": batch, "error": str(exc)})
    return {"selected": len(unique_paths), "deleted": deleted, "failed": failed}


def delete_report_rows(report_ids: list[str], execute: bool) -> dict:
    if not report_ids:
        return {"selected": 0, "deleted": 0}
    if not execute:
        return {"selected": len(report_ids), "deleted": 0}
    result = (
        app.get_supabase()
        .table(app.REPORTS_TABLE)
        .delete()
        .in_("id", report_ids)
        .execute()
    )
    return {"selected": len(report_ids), "deleted": len(result.data or [])}


def refund_submission_counts(duplicates: list[dict], execute: bool) -> list[dict]:
    refunds: dict[str, int] = defaultdict(int)
    for report in duplicates:
        if report.get("status") in COUNTED_STATUSES:
            refunds[report.get("user_id") or ""] += 1

    changes = []
    for user_id, refund in sorted(refunds.items()):
        if not user_id or refund <= 0:
            continue
        user = app.find_user_by_id(user_id)
        if not user:
            changes.append({"user_id": user_id, "refund": refund, "error": "user not found"})
            continue
        before = int(user.get("submissions_used") or 0)
        after = max(0, before - refund)
        change = {
            "user_id": user_id,
            "email": user.get("email", ""),
            "refund": refund,
            "before": before,
            "after": after,
        }
        if execute and after != before:
            (
                app.get_supabase()
                .table(app.SUPABASE_TABLE)
                .update({"submissions_used": after})
                .eq("id", user_id)
                .execute()
            )
        changes.append(change)
    return changes


def build_plan() -> dict:
    reports = app.list_reports_for_admin()
    groups = duplicate_groups(reports)
    duplicates = [report for group in groups for report in group["duplicates"]]
    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total_reports": len(reports),
        "duplicate_groups": len(groups),
        "duplicate_records": len(duplicates),
        "groups": groups,
    }


def run_dedupe(execute: bool, backup_dir: Path) -> dict:
    backup_dir.mkdir(parents=True, exist_ok=True)
    plan = build_plan()
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = backup_dir / f"duplicate-report-cleanup-{timestamp}.json"
    backup_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    duplicates = [report for group in plan["groups"] for report in group["duplicates"]]
    report_ids = [report["id"] for report in duplicates]
    storage_paths = [path for report in duplicates for path in storage_paths_for_delete(report)]

    storage_result = delete_supabase_storage(storage_paths, execute)
    row_result = delete_report_rows(report_ids, execute)
    refunds = refund_submission_counts(duplicates, execute)

    summary = {
        "execute": execute,
        "backup_path": str(backup_path),
        "duplicate_groups": plan["duplicate_groups"],
        "duplicate_records": plan["duplicate_records"],
        "storage": storage_result,
        "rows": row_result,
        "refunds": refunds,
    }
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Remove duplicate thesis audit records with the same user and original SHA256.")
    parser.add_argument("--execute", action="store_true", help="Actually delete duplicate rows and Supabase Storage objects.")
    parser.add_argument("--backup-dir", default="maintenance_backups", help="Directory for JSON cleanup plans.")
    args = parser.parse_args()

    summary = run_dedupe(args.execute, Path(args.backup_dir))
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

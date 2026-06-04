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


def archive_size(report: dict) -> int:
    return int(report.get("original_size_bytes") or 0) + int(report.get("report_size_bytes") or 0)


def build_prune_plan(keep_latest: int) -> dict:
    keep_count = max(1, int(keep_latest or 1))
    reports = app.list_reports_for_admin()
    by_user: dict[str, list[dict]] = defaultdict(list)
    for report in reports:
        user_id = report.get("user_id") or ""
        if user_id:
            by_user[user_id].append(report)

    groups = []
    stale_reports = []
    for user_id, rows in sorted(by_user.items()):
        sorted_rows = sorted(rows, key=lambda item: item.get("created_at") or "", reverse=True)
        stale = [row for row in sorted_rows[keep_count:] if app.report_has_any_archive(row)]
        if not stale:
            continue
        stale_reports.extend(stale)
        groups.append(
            {
                "user_id": user_id,
                "user_email": sorted_rows[0].get("user_email") or "",
                "kept_count": min(len(sorted_rows), keep_count),
                "stale_count": len(stale),
                "stale_archive_bytes": sum(archive_size(row) for row in stale),
                "stale_reports": [
                    {
                        "id": row.get("id"),
                        "created_at": row.get("created_at"),
                        "status": row.get("status") or "",
                        "original_filename": row.get("original_filename") or "",
                        "report_filename": row.get("report_filename") or "",
                        "original_storage_path": row.get("original_storage_path") or "",
                        "original_gcs_path": row.get("original_gcs_path") or "",
                        "original_drive_file_id": row.get("original_drive_file_id") or "",
                        "original_size_bytes": int(row.get("original_size_bytes") or 0),
                        "report_storage_path": row.get("report_storage_path") or "",
                        "report_gcs_path": row.get("report_gcs_path") or "",
                        "report_size_bytes": int(row.get("report_size_bytes") or 0),
                    }
                    for row in stale
                ],
            }
        )

    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "keep_latest": keep_count,
        "total_reports": len(reports),
        "affected_users": len(groups),
        "stale_reports": len(stale_reports),
        "stale_archive_bytes": sum(archive_size(row) for row in stale_reports),
        "groups": groups,
    }


def run_prune(execute: bool, keep_latest: int, backup_dir: Path) -> dict:
    backup_dir.mkdir(parents=True, exist_ok=True)
    plan = build_prune_plan(keep_latest)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = backup_dir / f"report-archive-prune-{timestamp}.json"
    backup_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    stale_reports = [row for group in plan["groups"] for row in group["stale_reports"]]
    if execute:
        delete_result = app.delete_archives_for_reports(stale_reports, clear_rows=True)
    else:
        original_targets = app.original_archive_delete_targets(stale_reports)
        report_targets = app.report_archive_delete_targets(stale_reports)
        delete_result = {
            "reports": len(stale_reports),
            "supabase": {"selected": len(set(original_targets["supabase_paths"] + report_targets["supabase_paths"])), "deleted": 0, "failed": []},
            "gcs_originals": {"selected": len(original_targets["gcs_paths"]), "deleted": 0, "failed": []},
            "gcs_reports": {"selected": len(report_targets["gcs_paths"]), "deleted": 0, "failed": []},
            "drive": {"selected": len(original_targets["drive_file_ids"]), "deleted": 0, "failed": []},
            "rows_cleared": 0,
        }

    return {
        "execute": execute,
        "backup_path": str(backup_path),
        "keep_latest": plan["keep_latest"],
        "affected_users": plan["affected_users"],
        "stale_reports": plan["stale_reports"],
        "stale_archive_bytes": plan["stale_archive_bytes"],
        "stale_archive_gb": round(plan["stale_archive_bytes"] / (1024 ** 3), 3),
        "delete": delete_result,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Keep only the newest downloadable report archives for each account.")
    parser.add_argument("--execute", action="store_true", help="Actually delete stale files and clear old download fields.")
    parser.add_argument("--keep-latest", type=int, default=app.MAX_STORED_REPORTS_PER_USER, help="How many newest reports to keep per account.")
    parser.add_argument("--backup-dir", default="maintenance_backups", help="Directory for JSON cleanup plans.")
    args = parser.parse_args()

    summary = run_prune(args.execute, args.keep_latest, Path(args.backup_dir))
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

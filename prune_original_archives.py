from __future__ import annotations

import argparse
import json
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import app


def report_has_original_archive(report: dict) -> bool:
    return bool(
        report.get("original_storage_path")
        or report.get("original_gcs_path")
        or report.get("original_drive_file_id")
        or int(report.get("original_size_bytes") or 0) > 0
        or report.get("original_sha256")
    )


def build_prune_plan() -> dict:
    reports = app.list_reports_for_admin()
    by_user: dict[str, list[dict]] = defaultdict(list)
    for report in reports:
        by_user[report.get("user_id") or ""].append(report)

    groups = []
    stale_reports = []
    for user_id, rows in sorted(by_user.items()):
        if not user_id:
            continue
        sorted_rows = sorted(rows, key=lambda item: item.get("created_at") or "", reverse=True)
        keep = sorted_rows[0] if sorted_rows else None
        stale = [row for row in sorted_rows[1:] if report_has_original_archive(row)]
        if not keep or not stale:
            continue
        stale_reports.extend(stale)
        groups.append(
            {
                "user_id": user_id,
                "user_email": keep.get("user_email") or "",
                "keep_report_id": keep.get("id"),
                "keep_created_at": keep.get("created_at"),
                "keep_original_filename": keep.get("original_filename") or "",
                "stale_count": len(stale),
                "stale_original_bytes": sum(int(row.get("original_size_bytes") or 0) for row in stale),
                "stale_reports": [
                    {
                        "id": row.get("id"),
                        "created_at": row.get("created_at"),
                        "original_filename": row.get("original_filename") or "",
                        "original_storage_path": row.get("original_storage_path") or "",
                        "original_gcs_path": row.get("original_gcs_path") or "",
                        "original_drive_file_id": row.get("original_drive_file_id") or "",
                        "original_size_bytes": int(row.get("original_size_bytes") or 0),
                        "original_sha256": row.get("original_sha256") or "",
                    }
                    for row in stale
                ],
            }
        )

    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total_reports": len(reports),
        "affected_users": len(groups),
        "stale_reports": len(stale_reports),
        "stale_original_bytes": sum(int(row.get("original_size_bytes") or 0) for row in stale_reports),
        "groups": groups,
        "stale_report_ids": [row["id"] for row in stale_reports if row.get("id")],
    }


def run_prune(execute: bool, backup_dir: Path) -> dict:
    backup_dir.mkdir(parents=True, exist_ok=True)
    plan = build_prune_plan()
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = backup_dir / f"original-archive-prune-{timestamp}.json"
    backup_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    stale_reports = [
        row
        for group in plan["groups"]
        for row in group["stale_reports"]
    ]
    if execute:
        delete_result = app.delete_original_archives_for_reports(stale_reports, clear_rows=True)
    else:
        targets = app.original_archive_delete_targets(stale_reports)
        delete_result = {
            "reports": len(stale_reports),
            "supabase": {"selected": len(targets["supabase_paths"]), "deleted": 0, "failed": []},
            "gcs": {"selected": len(targets["gcs_paths"]), "deleted": 0, "failed": []},
            "drive": {"selected": len(targets["drive_file_ids"]), "deleted": 0, "failed": []},
            "rows_cleared": 0,
        }

    return {
        "execute": execute,
        "backup_path": str(backup_path),
        "affected_users": plan["affected_users"],
        "stale_reports": plan["stale_reports"],
        "stale_original_bytes": plan["stale_original_bytes"],
        "stale_original_gb": round(plan["stale_original_bytes"] / (1024 ** 3), 3),
        "delete": delete_result,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Keep only the newest original upload archive for each account.")
    parser.add_argument("--execute", action="store_true", help="Actually delete stale original files and clear old original archive fields.")
    parser.add_argument("--backup-dir", default="maintenance_backups", help="Directory for JSON cleanup plans.")
    args = parser.parse_args()

    summary = run_prune(args.execute, Path(args.backup_dir))
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

from __future__ import annotations

import argparse
import json
import math
import os
import sys
from collections import defaultdict, deque
from datetime import datetime, timezone
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import app


def original_supabase_bytes(report: dict) -> int:
    if not report.get("original_storage_path"):
        return 0
    return int(report.get("original_size_bytes") or 0)


def build_fair_cleanup_selection(candidates: list[dict], target_bytes: int) -> list[dict]:
    by_user: dict[str, deque[dict]] = defaultdict(deque)
    for report in sorted(candidates, key=lambda item: item.get("created_at") or ""):
        user_key = report.get("user_id") or report.get("user_email") or "unknown"
        by_user[user_key].append(report)

    user_queue = deque(sorted(by_user))
    selected = []
    selected_bytes = 0
    while user_queue and selected_bytes < target_bytes:
        user_key = user_queue.popleft()
        user_reports = by_user[user_key]
        if not user_reports:
            continue
        report = user_reports.popleft()
        selected.append(report)
        selected_bytes += original_supabase_bytes(report)
        if user_reports:
            user_queue.append(user_key)
    return selected


def build_cleanup_plan(
    *,
    capacity_gb: float,
    trigger_ratio: float,
    delete_fraction: float,
    force: bool,
) -> dict:
    reports = app.list_reports_for_admin()
    supabase_original_reports = [report for report in reports if report.get("original_storage_path")]
    current_bytes = sum(original_supabase_bytes(report) for report in supabase_original_reports)
    capacity_bytes = int(capacity_gb * (1024 ** 3))
    trigger_bytes = int(capacity_bytes * trigger_ratio)
    usage_ratio = current_bytes / capacity_bytes if capacity_bytes else 0
    should_cleanup = force or current_bytes >= trigger_bytes

    candidates = [
        report
        for report in supabase_original_reports
        if report.get("original_drive_file_id")
    ]
    target_bytes = math.ceil(current_bytes * delete_fraction) if should_cleanup else 0
    selected = build_fair_cleanup_selection(candidates, target_bytes)
    selected_bytes = sum(original_supabase_bytes(report) for report in selected)

    by_user_summary: dict[str, dict] = {}
    for report in selected:
        user_key = report.get("user_id") or report.get("user_email") or "unknown"
        summary = by_user_summary.setdefault(
            user_key,
            {
                "user_id": report.get("user_id") or "",
                "user_email": report.get("user_email") or "",
                "reports": 0,
                "bytes": 0,
            },
        )
        summary["reports"] += 1
        summary["bytes"] += original_supabase_bytes(report)

    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "capacity_gb": capacity_gb,
        "trigger_ratio": trigger_ratio,
        "delete_fraction": delete_fraction,
        "force": force,
        "current_supabase_original_bytes": current_bytes,
        "current_supabase_original_gb": round(current_bytes / (1024 ** 3), 3),
        "usage_ratio": round(usage_ratio, 4),
        "trigger_bytes": trigger_bytes,
        "trigger_gb": round(trigger_bytes / (1024 ** 3), 3),
        "should_cleanup": should_cleanup,
        "eligible_reports": len(candidates),
        "selected_reports": len(selected),
        "selected_bytes": selected_bytes,
        "selected_gb": round(selected_bytes / (1024 ** 3), 3),
        "estimated_remaining_supabase_original_gb": round(max(current_bytes - selected_bytes, 0) / (1024 ** 3), 3),
        "affected_users": len(by_user_summary),
        "by_user": sorted(by_user_summary.values(), key=lambda item: (-item["bytes"], item["user_email"])),
        "items": [
            {
                "id": report.get("id"),
                "user_id": report.get("user_id") or "",
                "user_email": report.get("user_email") or "",
                "created_at": report.get("created_at") or "",
                "original_filename": report.get("original_filename") or "",
                "original_storage_path": report.get("original_storage_path") or "",
                "original_drive_file_id": report.get("original_drive_file_id") or "",
                "original_size_bytes": original_supabase_bytes(report),
            }
            for report in selected
        ],
    }


def execute_cleanup(selected_items: list[dict]) -> dict:
    paths_by_report_id = {
        item["id"]: item["original_storage_path"]
        for item in selected_items
        if item.get("id") and item.get("original_storage_path")
    }
    delete_result = app.delete_supabase_storage_objects(list(paths_by_report_id.values()))
    failed_paths = {
        path
        for failure in delete_result.get("failed", [])
        for path in failure.get("paths", [])
    }
    cleared_report_ids = [
        report_id
        for report_id, path in paths_by_report_id.items()
        if path not in failed_paths
    ]
    return {
        "supabase": delete_result,
        "rows_cleared": app.clear_supabase_original_archive_fields(cleared_report_ids),
        "rows_requested": len(cleared_report_ids),
    }


def run_maintenance(
    *,
    execute: bool,
    backup_dir: Path,
    capacity_gb: float,
    trigger_ratio: float,
    delete_fraction: float,
    force: bool,
) -> dict:
    backup_dir.mkdir(parents=True, exist_ok=True)
    plan = build_cleanup_plan(
        capacity_gb=capacity_gb,
        trigger_ratio=trigger_ratio,
        delete_fraction=delete_fraction,
        force=force,
    )
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    plan_path = backup_dir / f"supabase-original-storage-maintenance-{timestamp}.json"
    plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    if execute and plan["should_cleanup"] and plan["selected_reports"]:
        cleanup = execute_cleanup(plan["items"])
    else:
        cleanup = {
            "supabase": {"selected": len({item["original_storage_path"] for item in plan["items"]}), "deleted": 0, "failed": []},
            "rows_cleared": 0,
            "rows_requested": 0,
        }

    return {
        "execute": execute,
        "plan_path": str(plan_path),
        **{key: value for key, value in plan.items() if key != "items"},
        "cleanup": cleanup,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Free Supabase original-file storage after originals are backed up to Google Drive."
    )
    parser.add_argument("--execute", action="store_true", help="Actually delete Supabase original copies and clear Supabase path fields.")
    parser.add_argument("--backup-dir", default="maintenance_backups", help="Directory for JSON cleanup plans.")
    parser.add_argument("--capacity-gb", type=float, default=float(os.environ.get("SUPABASE_STORAGE_CAPACITY_GB", "1")), help="Supabase storage capacity used for the threshold calculation.")
    parser.add_argument("--trigger-ratio", type=float, default=float(os.environ.get("SUPABASE_STORAGE_TRIGGER_RATIO", "0.9")), help="Cleanup trigger ratio. Default: 0.9.")
    parser.add_argument("--delete-fraction", type=float, default=float(os.environ.get("SUPABASE_STORAGE_DELETE_FRACTION", "0.5")), help="Fraction of current Supabase original bytes to remove. Default: 0.5.")
    parser.add_argument("--force", action="store_true", help="Build and execute the deletion plan even below the trigger ratio.")
    args = parser.parse_args()

    if args.capacity_gb <= 0:
        parser.error("--capacity-gb must be greater than 0.")
    if not 0 < args.trigger_ratio <= 1:
        parser.error("--trigger-ratio must be in (0, 1].")
    if not 0 < args.delete_fraction <= 1:
        parser.error("--delete-fraction must be in (0, 1].")

    summary = run_maintenance(
        execute=args.execute,
        backup_dir=Path(args.backup_dir),
        capacity_gb=args.capacity_gb,
        trigger_ratio=args.trigger_ratio,
        delete_fraction=args.delete_fraction,
        force=args.force,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

import app


SCOPES = ["https://www.googleapis.com/auth/drive.file"]
DEFAULT_CLIENT_SECRET = os.environ.get("GOOGLE_DRIVE_OAUTH_CLIENT_SECRET", "")


def load_drive_service(client_secret: Path, token_path: Path):
    credentials = None
    if token_path.exists():
        credentials = Credentials.from_authorized_user_info(json.loads(token_path.read_text(encoding="utf-8")), SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(client_secret), SCOPES)
            credentials = flow.run_local_server(port=0, prompt="consent")
        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(credentials.to_json(), encoding="utf-8")
        token_path.chmod(0o600)
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def find_or_create_folder(service, name: str, parent_id: str | None = None) -> str:
    escaped_name = name.replace("\\", "\\\\").replace("'", "\\'")
    clauses = [
        "mimeType='application/vnd.google-apps.folder'",
        f"name='{escaped_name}'",
        "trashed=false",
    ]
    if parent_id:
        clauses.append(f"'{parent_id}' in parents")
    result = service.files().list(q=" and ".join(clauses), fields="files(id,name)", pageSize=1).execute()
    files = result.get("files", [])
    if files:
        return files[0]["id"]
    metadata = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    if parent_id:
        metadata["parents"] = [parent_id]
    created = service.files().create(body=metadata, fields="id").execute()
    return created["id"]


def ensure_folder_path(service, root_folder_name: str, parts: list[str]) -> str:
    current = find_or_create_folder(service, root_folder_name)
    for part in parts:
        clean = (part or "").strip()[:120] or "unknown"
        current = find_or_create_folder(service, clean, current)
    return current


def upload_bytes(service, folder_id: str, filename: str, content: bytes, content_type: str, description: str) -> str:
    metadata = {
        "name": filename[:180] or "archive.bin",
        "parents": [folder_id],
        "description": description[:1200],
    }
    media = MediaIoBaseUpload(app.io.BytesIO(content), mimetype=content_type, resumable=False)
    created = service.files().create(body=metadata, media_body=media, fields="id").execute()
    return created["id"]


def report_needs_migration(report: dict, include_reports: bool, include_originals: bool) -> bool:
    needs_original = include_originals and not report.get("original_drive_file_id") and (
        report.get("original_storage_path") or report.get("original_gcs_path")
    )
    needs_report = include_reports and (report.get("report_storage_path") or report.get("report_gcs_path"))
    return bool(needs_original or needs_report)


def update_original_drive_fields(report_id: str, file_id: str, drive_path: str) -> None:
    (
        app.get_supabase()
        .table(app.REPORTS_TABLE)
        .update({"original_drive_file_id": file_id, "original_drive_path": drive_path})
        .eq("id", report_id)
        .execute()
    )


def migrate_archives(
    *,
    execute: bool,
    client_secret: Path,
    token_path: Path,
    manifest_dir: Path,
    root_folder_name: str,
    include_reports: bool,
    include_originals: bool,
    limit: int,
) -> dict:
    reports = [item for item in app.list_reports_for_admin() if report_needs_migration(item, include_reports, include_originals)]
    if limit > 0:
        reports = reports[:limit]
    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "execute": execute,
        "root_folder_name": root_folder_name,
        "selected_reports": len(reports),
        "uploaded_originals": 0,
        "uploaded_reports": 0,
        "skipped": [],
        "failed": [],
        "items": [],
    }
    if not execute:
        summary["items"] = [
            {
                "report_id": item.get("id"),
                "user_email": item.get("user_email") or "",
                "created_at": item.get("created_at") or "",
                "original_filename": item.get("original_filename") or "",
                "needs_original": bool(include_originals and not item.get("original_drive_file_id") and (item.get("original_storage_path") or item.get("original_gcs_path"))),
                "needs_report": bool(include_reports and (item.get("report_storage_path") or item.get("report_gcs_path"))),
            }
            for item in reports
        ]
        return summary

    service = load_drive_service(client_secret, token_path)
    for report in reports:
        report_id = report.get("id") or ""
        user_email = report.get("user_email") or "unknown"
        created_at = (report.get("created_at") or "")[:10] or "unknown-date"
        folder_id = ensure_folder_path(service, root_folder_name, [user_email, created_at])
        item_log = {
            "report_id": report_id,
            "user_email": user_email,
            "created_at": report.get("created_at") or "",
            "original_filename": report.get("original_filename") or "",
            "report_filename": report.get("report_filename") or "",
            "original_drive_file_id": report.get("original_drive_file_id") or "",
            "report_drive_file_id": "",
        }
        try:
            if include_originals and not report.get("original_drive_file_id") and (report.get("original_storage_path") or report.get("original_gcs_path")):
                original_bytes = app.download_original_from_storage(report)
                filename = report.get("original_filename") or "thesis.docx"
                drive_file_id = upload_bytes(
                    service,
                    folder_id,
                    filename,
                    original_bytes,
                    app.original_content_type(filename),
                    f"UPC thesis audit original; report_id={report_id}",
                )
                drive_path = f"{root_folder_name}/{user_email}/{created_at}/{filename}"
                update_original_drive_fields(report_id, drive_file_id, drive_path)
                item_log["original_drive_file_id"] = drive_file_id
                item_log["original_drive_path"] = drive_path
                summary["uploaded_originals"] += 1
            if include_reports and (report.get("report_storage_path") or report.get("report_gcs_path")):
                report_bytes = app.download_report_from_storage(report)
                filename = report.get("report_filename") or "thesis_format_audit_report.html"
                drive_file_id = upload_bytes(
                    service,
                    folder_id,
                    filename,
                    report_bytes,
                    "text/html; charset=utf-8",
                    f"UPC thesis audit HTML report; report_id={report_id}",
                )
                item_log["report_drive_file_id"] = drive_file_id
                summary["uploaded_reports"] += 1
        except Exception as exc:
            summary["failed"].append({"report_id": report_id, "error": str(exc)})
        summary["items"].append(item_log)

    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Migrate current Supabase thesis audit archives to the user's Google Drive.")
    parser.add_argument("--execute", action="store_true", help="Actually upload files. Without this flag, only prints a migration plan.")
    parser.add_argument(
        "--client-secret",
        default=DEFAULT_CLIENT_SECRET,
        help="Google OAuth client_secret JSON path. Defaults to GOOGLE_DRIVE_OAUTH_CLIENT_SECRET.",
    )
    parser.add_argument("--token-path", default="maintenance_backups/google-drive-token.json", help="Local OAuth token cache path.")
    parser.add_argument("--manifest-dir", default="maintenance_backups", help="Directory for migration manifests.")
    parser.add_argument("--root-folder", default="UPC本科论文格式检测归档", help="Google Drive root folder name.")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of report records for testing. 0 means no limit.")
    parser.add_argument("--no-reports", action="store_true", help="Skip uploading HTML reports.")
    parser.add_argument("--no-originals", action="store_true", help="Skip uploading original Word files.")
    args = parser.parse_args()
    if not args.client_secret:
        parser.error("--client-secret is required unless GOOGLE_DRIVE_OAUTH_CLIENT_SECRET is set.")

    manifest_dir = Path(args.manifest_dir)
    manifest_dir.mkdir(parents=True, exist_ok=True)
    summary = migrate_archives(
        execute=args.execute,
        client_secret=Path(args.client_secret).expanduser(),
        token_path=Path(args.token_path).expanduser(),
        manifest_dir=manifest_dir,
        root_folder_name=args.root_folder,
        include_reports=not args.no_reports,
        include_originals=not args.no_originals,
        limit=args.limit,
    )
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    manifest_path = manifest_dir / f"google-drive-migration-{timestamp}.json"
    manifest_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({**summary, "manifest_path": str(manifest_path)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

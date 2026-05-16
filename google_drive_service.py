import os
import secrets
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import urlencode

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

from config import get_paths, read_json, write_json


GOOGLE_AUTH_BASE_URL = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_DRIVE_SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/userinfo.email",
    "openid",
]


def _env(name: str) -> str:
    return os.environ.get(name, "").strip()


def is_drive_configured() -> bool:
    return bool(_env("GOOGLE_CLIENT_ID") and _env("GOOGLE_CLIENT_SECRET") and _env("GOOGLE_OAUTH_REDIRECT_URI"))


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _state_path() -> Path:
    return get_paths().drive_state_file


def _read_state() -> Dict[str, Any]:
    return read_json(_state_path()) or {}


def _write_state(state: Dict[str, Any]) -> Dict[str, Any]:
    write_json(_state_path(), state)
    return state


def _invalidate_token(state: Dict[str, Any], *, error: str) -> None:
    state["connected"] = False
    state["email"] = None
    state["token"] = None
    state["last_error"] = error
    state["updated_at"] = _now_iso()
    _write_state(state)


def _oauth_client_config() -> Dict[str, Any]:
    client_id = _env("GOOGLE_CLIENT_ID")
    client_secret = _env("GOOGLE_CLIENT_SECRET")
    redirect_uri = _env("GOOGLE_OAUTH_REDIRECT_URI")
    if not (client_id and client_secret and redirect_uri):
        raise RuntimeError(
            "Google Drive OAuth is not configured. "
            "Set GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, and GOOGLE_OAUTH_REDIRECT_URI."
        )
    return {
        "web": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [redirect_uri],
        }
    }


def _credentials_from_state(state: Dict[str, Any]) -> Optional[Credentials]:
    token = state.get("token")
    if not isinstance(token, dict):
        return None

    try:
        credentials = Credentials.from_authorized_user_info(token, GOOGLE_DRIVE_SCOPES)
    except Exception:
        return None

    if credentials.expired and credentials.refresh_token:
        try:
            credentials.refresh(Request())
        except RefreshError:
            _invalidate_token(
                state,
                error="Google Drive session expired. Please connect again.",
            )
            return None
        except Exception as exc:
            _invalidate_token(state, error=f"Google Drive token refresh failed: {exc}")
            return None

        state["token"] = {
            "token": credentials.token,
            "refresh_token": credentials.refresh_token,
            "token_uri": credentials.token_uri,
            "client_id": credentials.client_id,
            "client_secret": credentials.client_secret,
            "scopes": credentials.scopes,
        }
        state["connected"] = True
        state["last_error"] = None
        state["updated_at"] = _now_iso()
        _write_state(state)
    return credentials


def get_drive_status() -> Dict[str, Any]:
    state = _read_state()
    credentials = _credentials_from_state(state) if state else None
    state = _read_state()
    connected = bool(credentials and credentials.valid)
    return {
        "configured": is_drive_configured(),
        "connected": connected,
        "email": state.get("email"),
        "last_error": state.get("last_error"),
        "updated_at": state.get("updated_at"),
    }


def create_oauth_start_url() -> Dict[str, str]:
    if not is_drive_configured():
        raise RuntimeError(
            "Google Drive OAuth is not configured. "
            "Set GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, and GOOGLE_OAUTH_REDIRECT_URI."
        )

    oauth_state = secrets.token_urlsafe(32)
    state = _read_state()
    state["oauth_state"] = oauth_state
    state["last_error"] = None
    state["updated_at"] = _now_iso()
    _write_state(state)

    params = {
        "client_id": _env("GOOGLE_CLIENT_ID"),
        "redirect_uri": _env("GOOGLE_OAUTH_REDIRECT_URI"),
        "response_type": "code",
        "scope": " ".join(GOOGLE_DRIVE_SCOPES),
        "access_type": "offline",
        "prompt": "consent",
        "state": oauth_state,
        "include_granted_scopes": "true",
    }
    return {"authorization_url": f"{GOOGLE_AUTH_BASE_URL}?{urlencode(params)}"}


def handle_oauth_callback(*, code: str, state: str) -> Dict[str, Any]:
    saved_state = _read_state()
    expected_state = str(saved_state.get("oauth_state", ""))
    if not expected_state or state != expected_state:
        raise RuntimeError("Invalid OAuth state. Please try connecting again.")

    flow = Flow.from_client_config(
        _oauth_client_config(),
        scopes=GOOGLE_DRIVE_SCOPES,
        redirect_uri=_env("GOOGLE_OAUTH_REDIRECT_URI"),
    )
    flow.fetch_token(code=code)
    credentials = flow.credentials

    service = build("oauth2", "v2", credentials=credentials, cache_discovery=False)
    user_info = service.userinfo().get().execute()
    email = str(user_info.get("email", "")).strip() or None

    new_state = {
        "connected": True,
        "email": email,
        "oauth_state": None,
        "token": {
            "token": credentials.token,
            "refresh_token": credentials.refresh_token,
            "token_uri": credentials.token_uri,
            "client_id": credentials.client_id,
            "client_secret": credentials.client_secret,
            "scopes": credentials.scopes,
        },
        "last_error": None,
        "updated_at": _now_iso(),
    }
    _write_state(new_state)
    return get_drive_status()


def disconnect_drive() -> Dict[str, Any]:
    state = _read_state()
    state["connected"] = False
    state["email"] = None
    state["oauth_state"] = None
    state["token"] = None
    state["last_error"] = None
    state["updated_at"] = _now_iso()
    _write_state(state)
    return get_drive_status()


def _get_drive_service() -> Any:
    if not is_drive_configured():
        raise RuntimeError("Google Drive OAuth is not configured on this server.")
    state = _read_state()
    credentials = _credentials_from_state(state)
    if not credentials or not credentials.valid:
        raise RuntimeError("Google Drive account is not connected. Please connect from the dashboard.")
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def _find_folder(service: Any, *, name: str, parent_id: Optional[str]) -> Optional[str]:
    escaped_name = name.replace("'", "\\'")
    query_parts = [
        "mimeType = 'application/vnd.google-apps.folder'",
        f"name = '{escaped_name}'",
        "trashed = false",
    ]
    if parent_id:
        query_parts.append(f"'{parent_id}' in parents")
    response = (
        service.files()
        .list(
            q=" and ".join(query_parts),
            spaces="drive",
            fields="files(id, name)",
            pageSize=1,
        )
        .execute()
    )
    files = response.get("files", [])
    if not files:
        return None
    return files[0]["id"]


def _ensure_folder(service: Any, *, name: str, parent_id: Optional[str]) -> str:
    existing_id = _find_folder(service, name=name, parent_id=parent_id)
    if existing_id:
        return existing_id

    body: Dict[str, Any] = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if parent_id:
        body["parents"] = [parent_id]
    created = service.files().create(body=body, fields="id").execute()
    return created["id"]


def _find_file(service: Any, *, name: str, parent_id: str) -> Optional[str]:
    escaped_name = name.replace("'", "\\'")
    query = (
        f"name = '{escaped_name}' and '{parent_id}' in parents and "
        "mimeType != 'application/vnd.google-apps.folder' and trashed = false"
    )
    response = (
        service.files()
        .list(q=query, spaces="drive", fields="files(id, name, webViewLink)", pageSize=1)
        .execute()
    )
    files = response.get("files", [])
    if not files:
        return None
    return files[0]["id"]


def _upsert_file(service: Any, *, local_path: Path, parent_id: str) -> Dict[str, Any]:
    file_name = local_path.name
    media = MediaFileUpload(
        str(local_path),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False,
    )
    existing_id = _find_file(service, name=file_name, parent_id=parent_id)
    if existing_id:
        updated = (
            service.files()
            .update(
                fileId=existing_id,
                media_body=media,
                fields="id, name, webViewLink",
            )
            .execute()
        )
        return {
            "file_name": file_name,
            "file_id": updated.get("id"),
            "web_view_link": updated.get("webViewLink"),
            "action": "updated",
            "success": True,
        }

    created = (
        service.files()
        .create(
            body={"name": file_name, "parents": [parent_id]},
            media_body=media,
            fields="id, name, webViewLink",
        )
        .execute()
    )
    return {
        "file_name": file_name,
        "file_id": created.get("id"),
        "web_view_link": created.get("webViewLink"),
        "action": "created",
        "success": True,
    }


def upload_month_files(*, year: str, month: str, monthly_path: Path, yearly_total_path: Path) -> Dict[str, Any]:
    service = _get_drive_service()
    root_id = _ensure_folder(service, name="AutoExpenses", parent_id=None)
    year_id = _ensure_folder(service, name=str(year), parent_id=root_id)
    month_id = _ensure_folder(service, name=str(month).zfill(2), parent_id=year_id)

    result: Dict[str, Any] = {
        "connected": True,
        "folder_path": f"AutoExpenses/{year}/{str(month).zfill(2)}",
        "monthly_file": None,
        "yearly_file": None,
    }

    result["monthly_file"] = _upsert_file(service, local_path=monthly_path, parent_id=month_id)
    # Keep one yearly workbook per year in AutoExpenses/{year}.
    result["yearly_file"] = _upsert_file(service, local_path=yearly_total_path, parent_id=year_id)
    return result

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


BASE_DIR = Path(__file__).resolve().parent
WORKSPACE_CONFIG_PATH = BASE_DIR / "google_workspace_config.json"
TOKENS_DIR = BASE_DIR / "tokens" / "google"

SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]


@dataclass
class WorkspaceConfig:
    credentials_file: Path
    token_store_dir: Path
    default_sheet_range: str = "Datos!A:Z"
    default_drive_folder_id: str = ""
    default_template_presentation_id: str = ""


@dataclass
class WorkspaceSession:
    user_key: str
    credentials: Credentials


def load_workspace_config(path: Path | None = None) -> WorkspaceConfig:
    config_path = path or WORKSPACE_CONFIG_PATH
    if not config_path.exists():
        raise FileNotFoundError(
            "No se encontro google_workspace_config.json. "
            "Crea una copia de google_workspace_config.example.json y complétala."
        )

    raw = json.loads(config_path.read_text(encoding="utf-8"))
    credentials_file = Path(raw["credentials_file"]).expanduser().resolve()
    token_store_dir = Path(raw.get("token_store_dir") or TOKENS_DIR).expanduser().resolve()
    token_store_dir.mkdir(parents=True, exist_ok=True)

    return WorkspaceConfig(
        credentials_file=credentials_file,
        token_store_dir=token_store_dir,
        default_sheet_range=raw.get("default_sheet_range", "Datos!A:Z"),
        default_drive_folder_id=raw.get("default_drive_folder_id", ""),
        default_template_presentation_id=raw.get("default_template_presentation_id", ""),
    )


def sanitize_user_key(user_key: str) -> str:
    safe = "".join(char if char.isalnum() or char in {"-", "_", "."} else "_" for char in user_key.strip().lower())
    return safe or "default"


def token_path_for_user(config: WorkspaceConfig, user_key: str) -> Path:
    return config.token_store_dir / f"{sanitize_user_key(user_key)}.json"


def load_user_credentials(config: WorkspaceConfig, user_key: str) -> Credentials | None:
    token_path = token_path_for_user(config, user_key)
    if not token_path.exists():
        return None
    credentials = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
        token_path.write_text(credentials.to_json(), encoding="utf-8")
    return credentials


def authorize_user(config: WorkspaceConfig, user_key: str, port: int = 0) -> WorkspaceSession:
    credentials = load_user_credentials(config, user_key)
    if credentials and credentials.valid:
        return WorkspaceSession(user_key=sanitize_user_key(user_key), credentials=credentials)

    if not config.credentials_file.exists():
        raise FileNotFoundError(
            f"No se encontro el archivo de credenciales OAuth: {config.credentials_file}"
        )

    flow = InstalledAppFlow.from_client_secrets_file(str(config.credentials_file), SCOPES)
    credentials = flow.run_local_server(port=port)
    token_path = token_path_for_user(config, user_key)
    token_path.write_text(credentials.to_json(), encoding="utf-8")
    return WorkspaceSession(user_key=sanitize_user_key(user_key), credentials=credentials)


def build_services(session: WorkspaceSession) -> dict[str, Any]:
    return {
        "drive": build("drive", "v3", credentials=session.credentials),
        "sheets": build("sheets", "v4", credentials=session.credentials),
        "slides": build("slides", "v1", credentials=session.credentials),
    }


def read_sheet_rows(
    sheets_service,
    spreadsheet_id: str,
    cell_range: str,
) -> list[list[str]]:
    response = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=cell_range)
        .execute()
    )
    return response.get("values", [])


def create_drive_presentation(
    slides_service,
    drive_service,
    title: str,
    folder_id: str = "",
) -> dict[str, str]:
    presentation = slides_service.presentations().create(body={"title": title}).execute()
    presentation_id = presentation["presentationId"]
    if folder_id:
        drive_service.files().update(
            fileId=presentation_id,
            addParents=folder_id,
            removeParents="root",
            fields="id, parents",
        ).execute()
    return {
        "presentation_id": presentation_id,
        "presentation_url": f"https://docs.google.com/presentation/d/{presentation_id}/edit",
    }


def create_basic_title_slide(
    slides_service,
    presentation_id: str,
    title: str,
    subtitle: str,
) -> None:
    requests = [
        {
            "createSlide": {
                "slideLayoutReference": {"predefinedLayout": "TITLE"},
            }
        }
    ]
    slides_service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": requests},
    ).execute()

    presentation = slides_service.presentations().get(presentationId=presentation_id).execute()
    slide = presentation["slides"][-1]
    title_id = None
    subtitle_id = None
    for element in slide.get("pageElements", []):
        placeholder = element.get("shape", {}).get("placeholder", {})
        if placeholder.get("type") == "TITLE":
            title_id = element["objectId"]
        elif placeholder.get("type") == "SUBTITLE":
            subtitle_id = element["objectId"]

    text_requests: list[dict[str, Any]] = []
    if title_id:
        text_requests.append({"insertText": {"objectId": title_id, "insertionIndex": 0, "text": title}})
    if subtitle_id:
        text_requests.append({"insertText": {"objectId": subtitle_id, "insertionIndex": 0, "text": subtitle}})
    if text_requests:
        slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": text_requests},
        ).execute()

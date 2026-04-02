"""
用 OAuth2 認證上傳 PPTX 至 Google Drive，自動轉換為 Google Slides
"""
import os
import pickle
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

BASE_DIR = Path("/home/barai/external_disk/barai/NKUST/project/114學年第二學期-協作教學")
CLIENT_SECRET = BASE_DIR / "app-token.json"
TOKEN_CACHE = BASE_DIR / "token.pickle"
PPTX_PATH = BASE_DIR / "產險分析" / "presentation_v2.pptx"

SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/presentations",
]


def get_credentials():
    creds = None
    if TOKEN_CACHE.exists():
        with open(TOKEN_CACHE, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRET), SCOPES)
            creds = flow.run_local_server(port=8099, open_browser=True)
        with open(TOKEN_CACHE, "wb") as f:
            pickle.dump(creds, f)

    return creds


def main():
    print("1. OAuth2 認證（如果是第一次，瀏覽器會跳出授權頁面）...")
    creds = get_credentials()
    print(f"   認證成功！")

    drive_service = build("drive", "v3", credentials=creds)

    print("2. 上傳 PPTX 並轉換為 Google Slides...")
    file_metadata = {
        "name": "產險資料分析 — 車險理賠預測實戰",
        "mimeType": "application/vnd.google-apps.presentation",
    }
    media = MediaFileUpload(
        str(PPTX_PATH),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        resumable=True,
    )
    file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,webViewLink",
    ).execute()

    file_id = file.get("id")
    web_link = file.get("webViewLink")

    print(f"   檔案 ID: {file_id}")

    url = f"https://docs.google.com/presentation/d/{file_id}/edit"
    print(f"\n{'='*60}")
    print(f"上傳完成！Google Slides 連結：")
    print(f"{url}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()

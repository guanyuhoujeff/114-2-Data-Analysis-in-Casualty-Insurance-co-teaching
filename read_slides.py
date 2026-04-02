"""讀取 Google Slides 簡報狀態"""
import pickle
from pathlib import Path
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

BASE_DIR = Path("/home/barai/external_disk/barai/NKUST/project/114學年第二學期-協作教學")
TOKEN_CACHE = BASE_DIR / "token.pickle"

def get_creds():
    with open(TOKEN_CACHE, "rb") as f:
        creds = pickle.load(f)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_CACHE, "wb") as f:
            pickle.dump(creds, f)
    return creds

def main():
    creds = get_creds()
    drive = build("drive", "v3", credentials=creds)

    # 找到剛上傳的簡報
    results = drive.files().list(
        q="mimeType='application/vnd.google-apps.presentation' and name contains '產險資料分析'",
        orderBy="modifiedTime desc",
        pageSize=5,
        fields="files(id,name,modifiedTime)"
    ).execute()

    files = results.get("files", [])
    if not files:
        print("找不到簡報！")
        return

    for f in files:
        print(f"  {f['name']} | ID: {f['id']} | {f['modifiedTime']}")

    pres_id = files[0]["id"]
    print(f"\n使用最新的: {pres_id}")

    # 讀取簡報結構
    slides_service = build("slides", "v1", credentials=creds)
    pres = slides_service.presentations().get(presentationId=pres_id).execute()

    print(f"\n簡報: {pres.get('title')}")
    print(f"投影片數: {len(pres.get('slides', []))}")
    print(f"尺寸: {pres.get('pageSize')}")

    # 列出每頁的元素
    for i, slide in enumerate(pres.get("slides", [])):
        slide_id = slide["objectId"]
        elements = slide.get("pageElements", [])

        # 找出文字內容
        texts = []
        for el in elements:
            shape = el.get("shape", {})
            tf = shape.get("text", {})
            for te in tf.get("textElements", []):
                tr = te.get("textRun", {})
                content = tr.get("content", "").strip()
                if content:
                    texts.append(content[:50])

        title = texts[0] if texts else "(空)"
        print(f"\n  [{i+1}] {slide_id}")
        print(f"      元素數: {len(elements)}")
        print(f"      文字: {title}")

    # 輸出 presentation ID 供後續使用
    with open(BASE_DIR / "產險分析" / "pres_id.txt", "w") as f:
        f.write(pres_id)
    print(f"\nPresentation ID 已存至 pres_id.txt")

if __name__ == "__main__":
    main()

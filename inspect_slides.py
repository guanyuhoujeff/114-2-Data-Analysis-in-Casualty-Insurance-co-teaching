"""深入檢查每頁投影片的元素位置與樣式"""
import json
import pickle
from pathlib import Path
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

BASE_DIR = Path("/home/barai/external_disk/barai/NKUST/project/114學年第二學期-協作教學")
TOKEN_CACHE = BASE_DIR / "token.pickle"
PRES_ID = (BASE_DIR / "產險分析" / "pres_id.txt").read_text().strip()

def get_creds():
    with open(TOKEN_CACHE, "rb") as f:
        creds = pickle.load(f)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return creds

def emu_to_inches(emu):
    return round(emu / 914400, 2)

def main():
    creds = get_creds()
    service = build("slides", "v1", credentials=creds)
    pres = service.presentations().get(presentationId=PRES_ID).execute()

    page_w = pres["pageSize"]["width"]["magnitude"]
    page_h = pres["pageSize"]["height"]["magnitude"]
    print(f"頁面尺寸: {emu_to_inches(page_w)}\" x {emu_to_inches(page_h)}\"")
    print(f"EMU: {page_w} x {page_h}\n")

    # 抽查幾頁：封面(0), 自介(1), 一般內容(8), section(7), 模型比較(24)
    check_pages = [0, 1, 2, 7, 8, 24, 34]

    for pg_idx in check_pages:
        slide = pres["slides"][pg_idx]
        slide_id = slide["objectId"]
        elements = slide.get("pageElements", [])
        print(f"{'='*60}")
        print(f"第 {pg_idx+1} 頁 ({slide_id}) — {len(elements)} 個元素")

        for el in elements:
            el_id = el.get("objectId", "?")
            size = el.get("size", {})
            transform = el.get("transform", {})

            w = size.get("width", {}).get("magnitude", 0)
            h = size.get("height", {}).get("magnitude", 0)
            tx = transform.get("translateX", 0)
            ty = transform.get("translateY", 0)
            sx = transform.get("scaleX", 1)
            sy = transform.get("scaleY", 1)

            # 取得文字
            shape = el.get("shape", {})
            tf = shape.get("text", {})
            texts = []
            font_info = ""
            for te in tf.get("textElements", []):
                tr = te.get("textRun", {})
                content = tr.get("content", "").strip()
                if content:
                    texts.append(content[:60])
                style = tr.get("style", {})
                if style.get("fontSize"):
                    font_info = f"fontSize={style['fontSize']}"
                if style.get("bold"):
                    font_info += " bold"

            text_preview = " | ".join(texts) if texts else "(no text)"

            # 計算實際位置
            actual_w = w * sx
            actual_h = h * sy

            print(f"  [{el_id}]")
            print(f"    位置: ({emu_to_inches(tx)}\", {emu_to_inches(ty)}\") "
                  f"大小: {emu_to_inches(actual_w)}\" x {emu_to_inches(actual_h)}\"")
            print(f"    {font_info}")
            print(f"    文字: {text_preview[:80]}")

    # 也存完整 JSON 供參考
    output = BASE_DIR / "產險分析" / "slides_dump.json"
    with open(output, "w", encoding="utf-8") as f:
        json.dump(pres, f, ensure_ascii=False, indent=2)
    print(f"\n完整 JSON 已存至 {output}")

if __name__ == "__main__":
    main()

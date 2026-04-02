"""
將 presentation_v2.html 解析後上傳至 Google Slides
使用 Service Account 認證，完成後分享給指定 Gmail
"""

import re
import html
from pathlib import Path
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── 設定 ──
CREDS_PATH = "/home/barai/external_disk/barai/NKUST/project/114學年第二學期-協作教學/credentials.json"
SHARE_EMAIL = "jeff7522553@gmail.com"
HTML_PATH = "/home/barai/external_disk/barai/NKUST/project/114學年第二學期-協作教學/產險分析/presentation_v2.html"
PRESENTATION_TITLE = "產險資料分析 — 車險理賠預測實戰"

SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]

# ── 色彩（Google Slides 用 0~1 RGB）──
def rgb(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}

ACCENT = rgb(0xD4, 0x48, 0x0B)
BLACK = rgb(0x1A, 0x1A, 0x1A)
GRAY = rgb(0x6B, 0x6B, 0x6B)
WHITE = rgb(1.0 * 255, 1.0 * 255, 1.0 * 255)  # will fix below
CREAM_BG = rgb(0xFF, 0xFD, 0xF7)
GREEN = rgb(0x1B, 0x7A, 0x3D)

WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
BLACK_SOLID = {"red": 0.1, "green": 0.1, "blue": 0.1}


def clean_text(el):
    if el is None:
        return ""
    txt = el.get_text(separator=" ", strip=True)
    txt = html.unescape(txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def pt(points):
    """Points to EMU"""
    return {"magnitude": points, "unit": "PT"}


def emu(value):
    return {"magnitude": value, "unit": "EMU"}


# 1 inch = 914400 EMU, slide is 10x5.625 inches (default 16:9)
SLIDE_W = 9144000  # 10 inches
SLIDE_H = 5143500  # 5.625 inches


def extract_slides(html_path):
    """解析 HTML 投影片"""
    with open(html_path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    slides = []
    for div in soup.find_all("div", class_="slide"):
        data = {
            "label": "",
            "title": "",
            "subtitle": "",
            "bullets": [],
            "cards": [],
            "code_blocks": [],
            "table_rows": [],
            "quote": "",
            "page_num": "",
            "is_section": False,
            "is_cover": False,
        }

        classes = div.get("class", [])

        page_el = div.find("span", class_="page-num")
        if page_el:
            data["page_num"] = clean_text(page_el)

        if "section-divider" in classes:
            data["is_section"] = True

        if "slide-center" in classes:
            h1 = div.find("h1")
            if h1 and ("產險" in clean_text(h1) or "概念" in clean_text(h1)):
                data["is_cover"] = True

        label_el = div.find("div", class_="label")
        if label_el:
            data["label"] = clean_text(label_el)

        h1 = div.find("h1")
        h2 = div.find("h2")
        if h1:
            data["title"] = clean_text(h1)
        elif h2:
            data["title"] = clean_text(h2)

        sub_p = div.find("p", class_="subtitle")
        if sub_p:
            data["subtitle"] = clean_text(sub_p)

        # Cards
        for card in div.find_all("div", class_=re.compile(r"card")):
            card_classes = card.get("class", [])
            card_type = "normal"
            if "card-accent" in card_classes:
                card_type = "accent"
            elif "card-green" in card_classes:
                card_type = "green"

            card_h3 = card.find("h3")
            card_ps = card.find_all("p")
            tag_el = card.find(class_=re.compile(r"tag"))

            card_data = {
                "type": card_type,
                "tag": clean_text(tag_el) if tag_el else "",
                "title": clean_text(card_h3) if card_h3 else "",
                "text": " ".join(clean_text(p) for p in card_ps if clean_text(p)),
            }
            if not card_data["title"] and not card_data["text"]:
                card_data["text"] = clean_text(card)
            data["cards"].append(card_data)

        # Tables
        table = div.find("table")
        if table:
            for tr in table.find_all("tr"):
                cells = [clean_text(td) for td in tr.find_all(["th", "td"])]
                if cells:
                    data["table_rows"].append(cells)

        # Code blocks
        for code in div.find_all("div", class_="code-block"):
            txt = clean_text(code)
            if txt:
                data["code_blocks"].append(txt)

        # Quote
        quote_el = div.find("div", class_="quote-block")
        if quote_el:
            data["quote"] = clean_text(quote_el)

        # Bullets / paragraphs
        for li in div.find_all("li"):
            if li.find_parent(class_=re.compile(r"card")):
                continue
            txt = clean_text(li)
            if txt:
                data["bullets"].append(txt)

        if not data["cards"] and not data["bullets"] and not data["table_rows"]:
            for p in div.find_all("p"):
                txt = clean_text(p)
                if (txt and txt != data["title"] and txt != data["subtitle"]
                    and txt != data["page_num"] and txt != data["label"]
                    and txt != data["quote"] and len(txt) > 3):
                    data["bullets"].append(txt)

        slides.append(data)
    return slides


def build_requests(slides_data):
    """生成 Google Slides API batch update requests"""
    requests = []

    for idx, data in enumerate(slides_data):
        slide_id = f"slide_{idx}"

        # ── 建立新投影片 ──
        if idx == 0:
            # 第一頁使用預設投影片（自動建立的）
            # 我們稍後取得它的 ID
            pass
        else:
            requests.append({
                "createSlide": {
                    "objectId": slide_id,
                    "insertionIndex": idx,
                    "slideLayoutReference": {"predefinedLayout": "BLANK"},
                }
            })

    return requests


def populate_slide(service, pres_id, slide_obj_id, data, idx):
    """填充單頁投影片內容"""
    requests = []

    is_dark = data["is_section"]

    # ── 背景色 ──
    if is_dark:
        bg_color = BLACK_SOLID
    else:
        bg_color = CREAM_BG

    requests.append({
        "updatePageProperties": {
            "objectId": slide_obj_id,
            "pageProperties": {
                "pageBackgroundFill": {
                    "solidFill": {"color": {"rgbColor": bg_color}}
                }
            },
            "fields": "pageBackgroundFill",
        }
    })

    elements = []  # (id, left, top, width, height, text, font_size, bold, color, alignment)

    if is_dark:
        # Section divider: label + title + subtitle (centered)
        if data["label"]:
            eid = f"s{idx}_label"
            elements.append({
                "id": eid, "left": 500000, "top": 1200000,
                "width": 8000000, "height": 400000,
                "text": data["label"], "font_size": 12,
                "bold": True, "color": ACCENT, "align": "CENTER",
            })
        eid = f"s{idx}_title"
        elements.append({
            "id": eid, "left": 500000, "top": 1600000,
            "width": 8000000, "height": 1500000,
            "text": data["title"], "font_size": 40,
            "bold": True, "color": WHITE, "align": "CENTER",
        })
        if data["subtitle"]:
            eid = f"s{idx}_sub"
            elements.append({
                "id": eid, "left": 500000, "top": 3200000,
                "width": 8000000, "height": 500000,
                "text": data["subtitle"], "font_size": 16,
                "bold": False, "color": {"red": 0.6, "green": 0.6, "blue": 0.6},
                "align": "CENTER",
            })
    elif data["is_cover"]:
        # Cover page
        if data["label"]:
            eid = f"s{idx}_label"
            elements.append({
                "id": eid, "left": 500000, "top": 800000,
                "width": 8000000, "height": 400000,
                "text": data["label"], "font_size": 11,
                "bold": True, "color": ACCENT, "align": "CENTER",
            })
        eid = f"s{idx}_title"
        elements.append({
            "id": eid, "left": 500000, "top": 1200000,
            "width": 8000000, "height": 1200000,
            "text": data["title"], "font_size": 42,
            "bold": True, "color": BLACK, "align": "CENTER",
        })
        if data["subtitle"]:
            eid = f"s{idx}_sub"
            elements.append({
                "id": eid, "left": 500000, "top": 2500000,
                "width": 8000000, "height": 500000,
                "text": data["subtitle"], "font_size": 20,
                "bold": False, "color": GRAY, "align": "CENTER",
            })
        # 底部資訊
        footer_text = ""
        for b in data["bullets"]:
            if "侯冠宇" in b or "2025" in b:
                footer_text = b
                break
        if footer_text:
            eid = f"s{idx}_footer"
            elements.append({
                "id": eid, "left": 500000, "top": 4200000,
                "width": 8000000, "height": 400000,
                "text": footer_text, "font_size": 10,
                "bold": False, "color": GRAY, "align": "CENTER",
            })
    else:
        # Normal slide
        y_cursor = 250000  # EMU from top

        if data["label"]:
            eid = f"s{idx}_label"
            elements.append({
                "id": eid, "left": 450000, "top": y_cursor,
                "width": 3000000, "height": 300000,
                "text": data["label"], "font_size": 10,
                "bold": True, "color": ACCENT, "align": "LEFT",
            })
            y_cursor += 300000

        if data["title"]:
            eid = f"s{idx}_title"
            elements.append({
                "id": eid, "left": 450000, "top": y_cursor,
                "width": 8200000, "height": 600000,
                "text": data["title"], "font_size": 26,
                "bold": True, "color": BLACK, "align": "LEFT",
            })
            y_cursor += 650000

        # 合併所有內容文字
        body_lines = []

        if data["subtitle"]:
            body_lines.append(data["subtitle"])

        # Table → 文字格式
        if data["table_rows"]:
            for row in data["table_rows"]:
                body_lines.append("  |  ".join(row))

        # Cards
        for card in data["cards"]:
            parts = []
            if card["tag"]:
                parts.append(f"[{card['tag']}]")
            if card["title"]:
                parts.append(card["title"])
            if card["text"]:
                parts.append(card["text"])
            if parts:
                body_lines.append(" — ".join(parts))

        # Code blocks
        for code in data["code_blocks"]:
            body_lines.append(f"Prompt: {code}")

        # Bullets
        for bullet in data["bullets"]:
            if data["is_cover"]:
                continue
            body_lines.append(f"• {bullet}")

        # Quote
        if data["quote"]:
            body_lines.append(f"「{data['quote']}」")

        if body_lines:
            body_text = "\n".join(body_lines)
            # 根據文字量調整字型大小
            font_size = 13
            if len(body_text) > 500:
                font_size = 11
            elif len(body_text) < 150:
                font_size = 15

            eid = f"s{idx}_body"
            remaining_height = 4800000 - y_cursor + 250000
            if remaining_height < 1000000:
                remaining_height = 3000000
            elements.append({
                "id": eid, "left": 450000, "top": y_cursor,
                "width": 8200000, "height": remaining_height,
                "text": body_text, "font_size": font_size,
                "bold": False, "color": GRAY, "align": "LEFT",
            })

    # 頁碼
    if data["page_num"]:
        eid = f"s{idx}_pn"
        pn_color = {"red": 0.6, "green": 0.6, "blue": 0.6} if is_dark else GRAY
        elements.append({
            "id": eid, "left": 7800000, "top": 4800000,
            "width": 1200000, "height": 300000,
            "text": data["page_num"], "font_size": 8,
            "bold": False, "color": pn_color, "align": "RIGHT",
        })

    # ── 生成 requests ──
    for el in elements:
        # 建立文字方塊
        requests.append({
            "createShape": {
                "objectId": el["id"],
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_obj_id,
                    "size": {
                        "width": emu(el["width"]),
                        "height": emu(el["height"]),
                    },
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": el["left"],
                        "translateY": el["top"],
                        "unit": "EMU",
                    },
                },
            }
        })
        # 插入文字
        requests.append({
            "insertText": {
                "objectId": el["id"],
                "text": el["text"],
                "insertionIndex": 0,
            }
        })
        # 設定文字樣式
        align_map = {
            "LEFT": "START",
            "CENTER": "CENTER",
            "RIGHT": "END",
        }
        requests.append({
            "updateTextStyle": {
                "objectId": el["id"],
                "style": {
                    "fontSize": pt(el["font_size"]),
                    "bold": el["bold"],
                    "foregroundColor": {
                        "opaqueColor": {"rgbColor": el["color"]}
                    },
                    "fontFamily": "Noto Sans TC",
                },
                "textRange": {"type": "ALL"},
                "fields": "fontSize,bold,foregroundColor,fontFamily",
            }
        })
        requests.append({
            "updateParagraphStyle": {
                "objectId": el["id"],
                "style": {
                    "alignment": align_map.get(el["align"], "START"),
                },
                "textRange": {"type": "ALL"},
                "fields": "alignment",
            }
        })

    if requests:
        service.presentations().batchUpdate(
            presentationId=pres_id,
            body={"requests": requests}
        ).execute()


def main():
    print("1. 解析 HTML...")
    slides_data = extract_slides(HTML_PATH)
    print(f"   共 {len(slides_data)} 頁")

    print("2. 認證 Google API...")
    creds = service_account.Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    slides_service = build("slides", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    print("3. 建立空白簡報...")
    pres = slides_service.presentations().create(
        body={"title": PRESENTATION_TITLE}
    ).execute()
    pres_id = pres["presentationId"]
    print(f"   ID: {pres_id}")

    # 取得第一頁的 ID
    first_slide_id = pres["slides"][0]["objectId"]

    print("4. 建立額外投影片...")
    create_requests = []
    slide_ids = [first_slide_id]
    for i in range(1, len(slides_data)):
        sid = f"slide_{i}"
        create_requests.append({
            "createSlide": {
                "objectId": sid,
                "insertionIndex": i,
                "slideLayoutReference": {"predefinedLayout": "BLANK"},
            }
        })
        slide_ids.append(sid)

    if create_requests:
        slides_service.presentations().batchUpdate(
            presentationId=pres_id,
            body={"requests": create_requests}
        ).execute()
    print(f"   建立了 {len(create_requests)} 頁空白投影片")

    print("5. 填充投影片內容...")
    for i, data in enumerate(slides_data):
        slide_obj_id = slide_ids[i]
        label = data["title"][:25] if data["title"] else "(empty)"
        print(f"   [{i+1}/{len(slides_data)}] {label}...")
        try:
            populate_slide(slides_service, pres_id, slide_obj_id, data, i)
        except Exception as e:
            print(f"   ⚠ 第 {i+1} 頁出錯: {e}")

    print("6. 分享給使用者...")
    drive_service.permissions().create(
        fileId=pres_id,
        body={
            "type": "user",
            "role": "writer",
            "emailAddress": SHARE_EMAIL,
        },
        sendNotificationEmail=True,
    ).execute()
    print(f"   已分享給 {SHARE_EMAIL}（編輯權限）")

    url = f"https://docs.google.com/presentation/d/{pres_id}/edit"
    print(f"\n{'='*60}")
    print(f"完成！Google Slides 連結：")
    print(f"{url}")
    print(f"{'='*60}")
    return url


if __name__ == "__main__":
    main()

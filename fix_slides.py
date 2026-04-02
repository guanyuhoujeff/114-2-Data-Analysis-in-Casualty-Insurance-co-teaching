"""
修正 Google Slides 排版：重新設定字型大小、元素位置與間距
策略：刪除所有舊元素，從 HTML 重新解析後以正確排版插入
"""
import re
import html as html_mod
import json
import pickle
from pathlib import Path
from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

BASE_DIR = Path("/home/barai/external_disk/barai/NKUST/project/114學年第二學期-協作教學")
TOKEN_CACHE = BASE_DIR / "token.pickle"
PRES_ID = (BASE_DIR / "產險分析" / "pres_id.txt").read_text().strip()
HTML_PATH = BASE_DIR / "產險分析" / "presentation_v2.html"

# ── 頁面尺寸 (EMU) ──
PAGE_W = 12191675  # 13.33"
PAGE_H = 6858000   # 7.5"

def inch(v):
    return int(v * 914400)

def pt_emu(v):
    return int(v * 12700)

def rgb(r, g, b):
    return {"red": r/255, "green": g/255, "blue": b/255}

ACCENT = rgb(0xD4, 0x48, 0x0B)
BLACK = rgb(0x1A, 0x1A, 0x1A)
GRAY = rgb(0x6B, 0x6B, 0x6B)
LIGHT_GRAY = rgb(0x99, 0x99, 0x99)
WHITE = rgb(255, 255, 255)
CREAM = rgb(0xFF, 0xFD, 0xF7)
DARK_BG = rgb(0x12, 0x12, 0x12)
GREEN = rgb(0x1B, 0x7A, 0x3D)
ACCENT_LIGHT = rgb(0xFF, 0xF0, 0xE8)
GREEN_LIGHT = rgb(0xE8, 0xF5, 0xEC)

def clean_text(el):
    if el is None:
        return ""
    txt = el.get_text(separator="\n", strip=True)
    txt = html_mod.unescape(txt)
    txt = re.sub(r"[ \t]+", " ", txt)
    # 保留換行
    lines = [l.strip() for l in txt.split("\n") if l.strip()]
    return "\n".join(lines)

def clean_inline(el):
    if el is None:
        return ""
    txt = el.get_text(separator=" ", strip=True)
    txt = html_mod.unescape(txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def extract_slides(html_path):
    with open(html_path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    slides = []
    for div in soup.find_all("div", class_="slide"):
        data = {
            "label": "", "title": "", "subtitle": "",
            "body_paragraphs": [],  # list of (text, style) where style = "normal"|"accent"|"green"|"code"|"quote"
            "table_rows": [],
            "page_num": "",
            "is_section": False, "is_cover": False,
        }
        classes = div.get("class", [])

        pn = div.find("span", class_="page-num")
        if pn:
            data["page_num"] = clean_inline(pn)

        if "section-divider" in classes:
            data["is_section"] = True

        if "slide-center" in classes:
            h1 = div.find("h1")
            if h1 and any(k in clean_inline(h1) for k in ["產險", "概念是你的"]):
                data["is_cover"] = True

        label = div.find("div", class_="label")
        if label:
            data["label"] = clean_inline(label)

        h1 = div.find("h1")
        h2 = div.find("h2")
        if h1:
            data["title"] = clean_inline(h1)
        elif h2:
            data["title"] = clean_inline(h2)

        sub = div.find("p", class_="subtitle")
        if sub:
            data["subtitle"] = clean_inline(sub)

        # Table
        table = div.find("table")
        if table:
            for tr in table.find_all("tr"):
                cells = [clean_inline(td) for td in tr.find_all(["th", "td"])]
                if cells:
                    data["table_rows"].append(cells)

        # Collect body content in order
        if not data["is_section"] and not data["is_cover"]:
            content_area = div.find("div", class_="slide-content") or div.find("div", class_="slide-split") or div

            # Walk through main children to preserve order
            collected_texts = set()

            # First, get direct paragraphs that aren't title/label/subtitle
            for p in div.find_all("p"):
                if p.find_parent(class_="page-num") or p.find_parent("span", class_="page-num"):
                    continue
                txt = clean_inline(p)
                if (txt and txt != data["title"] and txt != data["subtitle"]
                    and txt != data["page_num"] and txt != data["label"]
                    and len(txt) > 2 and txt not in collected_texts):
                    # Check if inside a card
                    parent_card = p.find_parent(class_=re.compile(r"card"))
                    if not parent_card:
                        data["body_paragraphs"].append((txt, "normal"))
                        collected_texts.add(txt)

            # Cards
            for card in div.find_all("div", class_=re.compile(r"card")):
                card_classes = card.get("class", [])
                style = "normal"
                if "card-accent" in card_classes:
                    style = "accent"
                elif "card-green" in card_classes:
                    style = "green"

                tag = card.find(class_=re.compile(r"tag"))
                h3 = card.find("h3")
                ps = card.find_all("p")

                parts = []
                if tag:
                    parts.append(f"[{clean_inline(tag)}]")
                if h3:
                    parts.append(clean_inline(h3))
                for cp in ps:
                    t = clean_inline(cp)
                    if t and t not in collected_texts:
                        parts.append(t)

                full = " — ".join(parts) if parts else clean_inline(card)
                if full and full not in collected_texts:
                    data["body_paragraphs"].append((full, style))
                    collected_texts.add(full)

            # Code blocks
            for code in div.find_all("div", class_="code-block"):
                txt = clean_inline(code)
                if txt and txt not in collected_texts:
                    data["body_paragraphs"].append((txt, "code"))
                    collected_texts.add(txt)

            # Quote
            quote = div.find("div", class_="quote-block")
            if quote:
                txt = clean_inline(quote)
                if txt and txt not in collected_texts:
                    data["body_paragraphs"].append((txt, "quote"))
                    collected_texts.add(txt)

            # Bullets (li)
            for li in div.find_all("li"):
                if li.find_parent(class_=re.compile(r"card")):
                    continue
                txt = clean_inline(li)
                if txt and txt not in collected_texts:
                    # Check for nested strong/code-block
                    strong = li.find("strong")
                    code = li.find("div", class_="code-block")
                    if strong and code:
                        data["body_paragraphs"].append((clean_inline(strong), "normal"))
                        data["body_paragraphs"].append((clean_inline(code), "code"))
                    else:
                        data["body_paragraphs"].append((f"• {txt}", "normal"))
                    collected_texts.add(txt)

        # For section/cover, also grab extra paragraphs
        if data["is_section"]:
            for p in div.find_all("p"):
                txt = clean_inline(p)
                if txt and txt != data["subtitle"] and txt != data["page_num"] and "rgba" not in str(p.get("style", "")):
                    if not data["subtitle"]:
                        data["subtitle"] = txt

        if data["is_cover"]:
            for p in div.find_all("p"):
                txt = clean_inline(p)
                if "侯冠宇" in txt or "2025" in txt:
                    data["subtitle"] = data.get("subtitle", "")
                    data["body_paragraphs"].append((txt, "normal"))
                    break

        slides.append(data)
    return slides


def get_creds():
    with open(TOKEN_CACHE, "rb") as f:
        creds = pickle.load(f)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return creds


def make_textbox(page_id, obj_id, left, top, width, height):
    return {
        "createShape": {
            "objectId": obj_id,
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": page_id,
                "size": {
                    "width": {"magnitude": width, "unit": "EMU"},
                    "height": {"magnitude": height, "unit": "EMU"},
                },
                "transform": {
                    "scaleX": 1, "scaleY": 1,
                    "translateX": left, "translateY": top,
                    "unit": "EMU",
                },
            },
        }
    }


def make_insert_text(obj_id, text):
    return {"insertText": {"objectId": obj_id, "text": text, "insertionIndex": 0}}


def make_text_style(obj_id, font_size, bold=False, color=BLACK, font="Noto Sans TC"):
    return {
        "updateTextStyle": {
            "objectId": obj_id,
            "style": {
                "fontSize": {"magnitude": font_size, "unit": "PT"},
                "bold": bold,
                "foregroundColor": {"opaqueColor": {"rgbColor": color}},
                "fontFamily": font,
            },
            "textRange": {"type": "ALL"},
            "fields": "fontSize,bold,foregroundColor,fontFamily",
        }
    }


def make_para_style(obj_id, alignment="START", space_above=0, space_below=0, line_spacing=115):
    style = {"alignment": alignment}
    fields = ["alignment"]
    if space_above:
        style["spaceAbove"] = {"magnitude": space_above, "unit": "PT"}
        fields.append("spaceAbove")
    if space_below:
        style["spaceBelow"] = {"magnitude": space_below, "unit": "PT"}
        fields.append("spaceBelow")
    if line_spacing:
        style["lineSpacing"] = line_spacing
        fields.append("lineSpacing")
    return {
        "updateParagraphStyle": {
            "objectId": obj_id,
            "style": style,
            "textRange": {"type": "ALL"},
            "fields": ",".join(fields),
        }
    }


def make_bg(page_id, color):
    return {
        "updatePageProperties": {
            "objectId": page_id,
            "pageProperties": {
                "pageBackgroundFill": {
                    "solidFill": {"color": {"rgbColor": color}}
                }
            },
            "fields": "pageBackgroundFill",
        }
    }


def make_rect_bg(page_id, obj_id, left, top, width, height, fill_color):
    return {
        "createShape": {
            "objectId": obj_id,
            "shapeType": "ROUND_RECTANGLE",
            "elementProperties": {
                "pageObjectId": page_id,
                "size": {
                    "width": {"magnitude": width, "unit": "EMU"},
                    "height": {"magnitude": height, "unit": "EMU"},
                },
                "transform": {
                    "scaleX": 1, "scaleY": 1,
                    "translateX": left, "translateY": top,
                    "unit": "EMU",
                },
            },
        }
    }


def batch_update(service, pres_id, requests):
    """分批發送 requests（每批最多 500 個）"""
    BATCH_SIZE = 400
    for i in range(0, len(requests), BATCH_SIZE):
        batch = requests[i:i+BATCH_SIZE]
        service.presentations().batchUpdate(
            presentationId=pres_id,
            body={"requests": batch}
        ).execute()


def main():
    print("1. 解析 HTML...")
    slides_data = extract_slides(HTML_PATH)
    print(f"   {len(slides_data)} 頁")

    print("2. 連接 Google Slides API...")
    creds = get_creds()
    service = build("slides", "v1", credentials=creds)

    print("3. 讀取現有簡報...")
    pres = service.presentations().get(presentationId=PRES_ID).execute()
    existing_slides = pres.get("slides", [])

    # ── Step 1: 刪除每頁的所有舊元素 ──
    print("4. 清除舊元素...")
    delete_requests = []
    for slide in existing_slides:
        for el in slide.get("pageElements", []):
            delete_requests.append({
                "deleteObject": {"objectId": el["objectId"]}
            })
    if delete_requests:
        batch_update(service, PRES_ID, delete_requests)
    print(f"   刪除了 {len(delete_requests)} 個舊元素")

    # ── Step 2: 逐頁重新填充 ──
    print("5. 重新排版...")

    MARGIN_L = inch(0.9)
    MARGIN_R = inch(0.9)
    CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R  # ~11.55"

    for i, data in enumerate(slides_data):
        page_id = existing_slides[i]["objectId"]
        reqs = []
        uid = 0

        def next_id():
            nonlocal uid
            uid += 1
            return f"fix_{i}_{uid}"

        label_text = data["title"][:30] if data["title"] else ""
        print(f"   [{i+1}/35] {label_text}...")

        if data["is_section"]:
            # ── Section Divider ──
            reqs.append(make_bg(page_id, DARK_BG))

            if data["label"]:
                oid = next_id()
                reqs.append(make_textbox(page_id, oid, MARGIN_L, inch(2.2), CONTENT_W, inch(0.4)))
                reqs.append(make_insert_text(oid, data["label"]))
                reqs.append(make_text_style(oid, 13, bold=True, color=ACCENT, font="Inter"))
                reqs.append(make_para_style(oid, "CENTER"))

            oid = next_id()
            title_size = 44 if len(data["title"]) < 15 else 36
            reqs.append(make_textbox(page_id, oid, MARGIN_L, inch(2.8), CONTENT_W, inch(1.8)))
            reqs.append(make_insert_text(oid, data["title"]))
            reqs.append(make_text_style(oid, title_size, bold=True, color=WHITE))
            reqs.append(make_para_style(oid, "CENTER", line_spacing=120))

            if data["subtitle"]:
                oid = next_id()
                reqs.append(make_textbox(page_id, oid, MARGIN_L, inch(4.8), CONTENT_W, inch(0.6)))
                reqs.append(make_insert_text(oid, data["subtitle"]))
                reqs.append(make_text_style(oid, 18, color=LIGHT_GRAY))
                reqs.append(make_para_style(oid, "CENTER"))

            # 頁碼
            if data["page_num"]:
                oid = next_id()
                reqs.append(make_textbox(page_id, oid, inch(11), inch(6.9), inch(2), inch(0.35)))
                reqs.append(make_insert_text(oid, data["page_num"]))
                reqs.append(make_text_style(oid, 9, color=LIGHT_GRAY, font="Inter"))
                reqs.append(make_para_style(oid, "END"))

        elif data["is_cover"]:
            # ── Cover ──
            reqs.append(make_bg(page_id, CREAM))

            if data["label"]:
                oid = next_id()
                reqs.append(make_textbox(page_id, oid, MARGIN_L, inch(1.8), CONTENT_W, inch(0.4)))
                reqs.append(make_insert_text(oid, data["label"]))
                reqs.append(make_text_style(oid, 12, bold=True, color=ACCENT, font="Inter"))
                reqs.append(make_para_style(oid, "CENTER"))

            oid = next_id()
            reqs.append(make_textbox(page_id, oid, MARGIN_L, inch(2.5), CONTENT_W, inch(1.4)))
            reqs.append(make_insert_text(oid, data["title"]))
            reqs.append(make_text_style(oid, 48, bold=True, color=BLACK))
            reqs.append(make_para_style(oid, "CENTER", line_spacing=110))

            if data["subtitle"]:
                oid = next_id()
                reqs.append(make_textbox(page_id, oid, MARGIN_L, inch(4.0), CONTENT_W, inch(0.6)))
                reqs.append(make_insert_text(oid, data["subtitle"]))
                reqs.append(make_text_style(oid, 22, color=GRAY))
                reqs.append(make_para_style(oid, "CENTER"))

            # Footer info
            for txt, style in data["body_paragraphs"]:
                if "侯冠宇" in txt or "2025" in txt:
                    oid = next_id()
                    reqs.append(make_textbox(page_id, oid, MARGIN_L, inch(5.5), CONTENT_W, inch(0.4)))
                    reqs.append(make_insert_text(oid, txt))
                    reqs.append(make_text_style(oid, 11, color=GRAY, font="Inter"))
                    reqs.append(make_para_style(oid, "CENTER"))
                    break

            if data["page_num"]:
                oid = next_id()
                reqs.append(make_textbox(page_id, oid, inch(11), inch(6.9), inch(2), inch(0.35)))
                reqs.append(make_insert_text(oid, data["page_num"]))
                reqs.append(make_text_style(oid, 9, color=GRAY, font="Inter"))
                reqs.append(make_para_style(oid, "END"))

        else:
            # ── Normal Slide ──
            reqs.append(make_bg(page_id, CREAM))
            y = inch(0.6)

            # Label
            if data["label"]:
                oid = next_id()
                reqs.append(make_textbox(page_id, oid, MARGIN_L, y, inch(4), inch(0.35)))
                reqs.append(make_insert_text(oid, data["label"]))
                reqs.append(make_text_style(oid, 11, bold=True, color=ACCENT, font="Inter"))
                reqs.append(make_para_style(oid, "START"))
                y += inch(0.4)

            # Title
            if data["title"]:
                oid = next_id()
                t_size = 30 if len(data["title"]) < 20 else 26
                reqs.append(make_textbox(page_id, oid, MARGIN_L, y, CONTENT_W, inch(0.7)))
                reqs.append(make_insert_text(oid, data["title"]))
                reqs.append(make_text_style(oid, t_size, bold=True, color=BLACK))
                reqs.append(make_para_style(oid, "START", line_spacing=110))
                y += inch(0.85)

            # Table
            if data["table_rows"]:
                rows = data["table_rows"]
                n_rows = len(rows)
                n_cols = max(len(r) for r in rows)
                row_h = inch(0.35)
                table_h = row_h * n_rows
                table_w = min(CONTENT_W, inch(11.5))

                reqs.append({
                    "createTable": {
                        "objectId": next_id(),
                        "elementProperties": {
                            "pageObjectId": page_id,
                            "size": {
                                "width": {"magnitude": table_w, "unit": "EMU"},
                                "height": {"magnitude": table_h, "unit": "EMU"},
                            },
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": MARGIN_L, "translateY": y,
                                "unit": "EMU",
                            },
                        },
                        "rows": n_rows,
                        "columns": n_cols,
                    }
                })
                # Note: table cell text needs separate requests referencing tableId
                # For simplicity, we'll add table as text below
                y += table_h + inch(0.3)

                # Render table as formatted text instead (more reliable)
                table_text = ""
                for ri, row in enumerate(rows):
                    table_text += "  |  ".join(row) + "\n"

                oid = next_id()
                reqs.append(make_textbox(page_id, oid, MARGIN_L, y - table_h - inch(0.1), CONTENT_W, table_h + inch(0.3)))
                reqs.append(make_insert_text(oid, table_text.strip()))
                reqs.append(make_text_style(oid, 12, color=BLACK, font="Noto Sans TC"))
                reqs.append(make_para_style(oid, "START", space_below=4, line_spacing=140))

            # Body paragraphs
            if data["body_paragraphs"]:
                # 計算剩餘空間
                remaining = PAGE_H - y - inch(0.8)
                if remaining < inch(1.5):
                    remaining = inch(4)

                # 根據內容量決定字型
                total_chars = sum(len(t) for t, s in data["body_paragraphs"])
                n_paras = len(data["body_paragraphs"])

                if total_chars > 600 or n_paras > 8:
                    body_font = 11
                    line_sp = 130
                elif total_chars > 300 or n_paras > 5:
                    body_font = 13
                    line_sp = 140
                else:
                    body_font = 15
                    line_sp = 150

                # 合併成一個文字方塊（多段落）
                full_text = "\n\n".join(t for t, s in data["body_paragraphs"])

                oid = next_id()
                reqs.append(make_textbox(page_id, oid, MARGIN_L, y, CONTENT_W, remaining))
                reqs.append(make_insert_text(oid, full_text))
                reqs.append(make_text_style(oid, body_font, color=GRAY))
                reqs.append(make_para_style(oid, "START", space_below=6, line_spacing=line_sp))

                # 針對特殊段落做部分樣式 (accent 的做 bold)
                # 用 range-based styling
                cursor = 0
                for txt, style in data["body_paragraphs"]:
                    start = full_text.find(txt, cursor)
                    if start == -1:
                        continue
                    end = start + len(txt)

                    if style == "accent":
                        reqs.append({
                            "updateTextStyle": {
                                "objectId": oid,
                                "style": {
                                    "foregroundColor": {"opaqueColor": {"rgbColor": ACCENT}},
                                    "bold": True,
                                },
                                "textRange": {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end},
                                "fields": "foregroundColor,bold",
                            }
                        })
                    elif style == "green":
                        reqs.append({
                            "updateTextStyle": {
                                "objectId": oid,
                                "style": {
                                    "foregroundColor": {"opaqueColor": {"rgbColor": GREEN}},
                                    "bold": True,
                                },
                                "textRange": {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end},
                                "fields": "foregroundColor,bold",
                            }
                        })
                    elif style == "code":
                        reqs.append({
                            "updateTextStyle": {
                                "objectId": oid,
                                "style": {
                                    "fontFamily": "Roboto Mono",
                                    "fontSize": {"magnitude": max(body_font - 2, 10), "unit": "PT"},
                                    "foregroundColor": {"opaqueColor": {"rgbColor": BLACK}},
                                    "backgroundColor": {"opaqueColor": {"rgbColor": rgb(0xF0, 0xF0, 0xF0)}},
                                },
                                "textRange": {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end},
                                "fields": "fontFamily,fontSize,foregroundColor,backgroundColor",
                            }
                        })
                    elif style == "quote":
                        reqs.append({
                            "updateTextStyle": {
                                "objectId": oid,
                                "style": {
                                    "foregroundColor": {"opaqueColor": {"rgbColor": ACCENT}},
                                    "italic": True,
                                    "fontSize": {"magnitude": body_font + 2, "unit": "PT"},
                                },
                                "textRange": {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end},
                                "fields": "foregroundColor,italic,fontSize",
                            }
                        })

                    cursor = end

            # 頁碼
            if data["page_num"]:
                oid = next_id()
                reqs.append(make_textbox(page_id, oid, inch(11), inch(6.9), inch(2), inch(0.35)))
                reqs.append(make_insert_text(oid, data["page_num"]))
                reqs.append(make_text_style(oid, 9, color=GRAY, font="Inter"))
                reqs.append(make_para_style(oid, "END"))

        # 發送這一頁的 requests
        if reqs:
            try:
                batch_update(service, PRES_ID, reqs)
            except Exception as e:
                print(f"      ⚠ 錯誤: {e}")

    url = f"https://docs.google.com/presentation/d/{PRES_ID}/edit"
    print(f"\n{'='*60}")
    print(f"排版修正完成！")
    print(f"{url}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()

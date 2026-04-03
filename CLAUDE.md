# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

產險資料分析（Data Analysis in Casualty Insurance）— NKUST 114-2 學期協同教學課程教材。
核心產出是一份 HTML 投影片（`presentation_v2.html`），搭配自動化 pipeline 匯出為 PPTX 及同步至 Google Slides。

課程主題：車險理賠預測（Car Insurance Claim Prediction），使用 Kaggle 10,000 筆資料集，涵蓋 EDA、分類模型（LR/Decision Tree/Random Forest/SVM/XGBoost）、精算洞察報告。教學法強調 Vibe Coding（AI 輔助自然語言寫程式）。

## Architecture

```
presentation_v2.html  (source of truth)
        │
        ├──► html_to_pptx.py        → presentation_v2.pptx (local PPTX)
        ├──► upload_to_google_slides.py → Google Slides (service account, API 直接建立)
        │
        └──► upload_oauth.py         → 上傳 PPTX 至 Google Drive (OAuth, 自動轉 Slides)
              
fix_slides.py      → 對已建立的 Google Slides 做格式修正（刪除舊元素、重建）
read_slides.py     → 查詢 Google Drive 取得 presentation ID → pres_id.txt
inspect_slides.py  → 匯出 slides_dump.json 供 debug 用
images/gen.py      → 使用 Gemini 2.5 Flash 生成圖片
```

## Key Commands

```bash
# HTML → PPTX 轉換
python html_to_pptx.py

# 直接透過 API 建立 Google Slides（需 credentials.json）
python upload_to_google_slides.py

# 上傳 PPTX 至 Google Drive（OAuth，首次需瀏覽器登入）
python upload_oauth.py

# 修正 Google Slides 格式
python fix_slides.py

# 查詢 presentation ID
python read_slides.py

# 檢查 slides 元素（debug）
python inspect_slides.py

# 生成圖片
python images/gen.py "prompt text" output.png
```

## Dependencies

python-pptx, beautifulsoup4, google-auth-oauthlib, google-api-python-client, google-auth-httplib2, google-genai, Pillow

## Design System Constants

所有 Python 腳本共用的色彩與尺寸定義：
- **ACCENT**: `#D4480B` (橘紅)
- **BLACK**: `#1A1A1A`
- **CREAM_BG**: `#FFFDF7` (投影片背景)
- **GREEN**: `#1B7A3D`
- Slide 尺寸：16:9（13.333" × 7.5"）
- 字體：Noto Sans TC（中文）、Inter（英文）

## Google API Notes

- `upload_to_google_slides.py` 使用 **Service Account**（`credentials.json`）
- `upload_oauth.py` / `fix_slides.py` / `read_slides.py` 使用 **OAuth2**（`token.pickle` 快取）
- `fix_slides.py` 每批次最多 400 個 API request，避免超過 Google API 限制
- Presentation ID 存於 `pres_id.txt`，供其他腳本引用

## Data

`Car_Insurance_Claim.csv`：10,000 筆車險理賠資料，19 features + 1 target（OUTCOME）。CREDIT_SCORE 和 ANNUAL_MILEAGE 有缺失值。

## Content Files

| 檔案 | 用途 |
|------|------|
| `lecture_plan.md` | 3 小時課程規劃（三節課結構） |
| `lecture_script.md` | 逐字講稿（含教學提示） |
| `slides_content.md` | 投影片文字內容大綱 |
| `student_handout.md` | 學生練習任務單（7 個 Task） |
| `lesson_meta.md` | 課程元資料（學分、評分方式、進度表） |

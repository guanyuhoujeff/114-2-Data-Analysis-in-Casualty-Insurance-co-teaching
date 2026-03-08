# 產險資料分析 — 學生練習任務清單

## 講座資訊

- 日期：4/9（四）5-7 節
- 工具：Google Antigravity（https://antigravity.google）
- 資料集：`Car_Insurance_Claim.csv`（Kaggle Car Insurance Claim Prediction）
- 程式語言：Python

---

## 資料集欄位說明

| 欄位 | 說明 | 型態 |
|------|------|------|
| `AGE` | 年齡區間（16-25 / 26-39 / 40-64 / 65+） | 類別 |
| `GENDER` | 性別（male / female） | 類別 |
| `RACE` | 種族（majority / minority） | 類別 |
| `DRIVING_EXPERIENCE` | 駕齡（0-9y / 10-19y / 20-29y / 30y+） | 類別 |
| `EDUCATION` | 教育程度（none / high school / university） | 類別 |
| `INCOME` | 收入階層（poverty / working class / middle class / upper class） | 類別 |
| `CREDIT_SCORE` | 信用分數（0~1，有缺失值） | 連續 |
| `VEHICLE_OWNERSHIP` | 車輛所有權（0: 租賃 / 1: 自有） | 二元 |
| `VEHICLE_YEAR` | 出廠年份（before 2015 / after 2015） | 類別 |
| `MARRIED` | 婚姻狀態（0 / 1） | 二元 |
| `CHILDREN` | 有無小孩（0 / 1） | 二元 |
| `POSTAL_CODE` | 郵遞區號 | 類別 |
| `ANNUAL_MILEAGE` | 年里程數（有缺失值） | 連續 |
| `VEHICLE_TYPE` | 車型（sedan / sports car） | 類別 |
| `SPEEDING_VIOLATIONS` | 超速違規次數 | 離散 |
| `DUIS` | 酒駕次數 | 離散 |
| `PAST_ACCIDENTS` | 過去事故次數 | 離散 |
| `OUTCOME` | 是否申請理賠（0: 否 / 1: 是）⬅ **預測目標** | 二元 |

---

## Part 1：資料探索（EDA）

### 任務 1：描述統計與驗證

**目標**：用 AI 產出描述統計，並手動驗證結果是否正確。

**步驟**：

1. 將 `Car_Insurance_Claim.csv` 上傳至 Antigravity
2. 輸入 Prompt，請 AI 計算所有數值變數的描述統計（平均數、中位數、標準差、最大值、最小值）
3. **驗證**：挑選一個變數（例如 `ANNUAL_MILEAGE`），用計算機或心算驗證平均數是否正確

**Prompt 參考**：
> 「請讀取 `Car_Insurance_Claim.csv`，計算所有數值欄位的描述統計（count、mean、std、min、25%、50%、75%、max），並以表格呈現。」

**思考題**：
- `CREDIT_SCORE` 的 count 是否等於總筆數？如果不是，代表什麼？
- `ANNUAL_MILEAGE` 的分佈是對稱的還是偏態的？你怎麼判斷？

---

### 任務 2：分佈視覺化與異常值觀察

**目標**：用圖表觀察資料分佈，找出可能的異常值。

**步驟**：

1. 請 AI 畫出 `ANNUAL_MILEAGE` 的直方圖（histogram）
2. 請 AI 畫出 `CREDIT_SCORE` 的箱形圖（boxplot）
3. 觀察圖表，回答以下問題

**Prompt 參考**：
> 「請畫出 `ANNUAL_MILEAGE` 的直方圖，設定 30 個 bins，並標示平均數的垂直線。」

> 「請畫出 `CREDIT_SCORE` 的箱形圖，標示出異常值。」

**思考題**：
- 直方圖的形狀是什麼？（左偏 / 右偏 / 常態）
- 箱形圖中有沒有超出鬍鬚範圍的點？這些點代表什麼？
- 如果有異常值，你會怎麼處理？刪除？還是保留？為什麼？

---

### 任務 3：交叉分析與高風險客群辨識

**目標**：透過多變數交叉分析，找出高風險客群。

**步驟**：

1. 自選兩個類別變數（例如 `AGE` × `DRIVING_EXPERIENCE`）
2. 請 AI 計算交叉分組下的理賠機率，並畫成熱力圖
3. 從熱力圖中辨識出理賠機率最高和最低的組合

**Prompt 參考**：
> 「請計算 `AGE` 與 `DRIVING_EXPERIENCE` 交叉分組下的理賠機率（`OUTCOME` 的平均值），並用 seaborn 畫成熱力圖，加上數值標籤。」

**思考題**：
- 哪個組合的理賠機率最高？這符合你的直覺嗎？
- 如果你是精算師，會如何利用這個結果調整保費？
- 試試換其他變數組合（例如 `INCOME` × `VEHICLE_TYPE`），結果有什麼不同？

---

## Part 2：分類模型

### 任務 4：Benchmark 模型 — Logistic Regression（講師帶做）

**目標**：建立 Logistic Regression 作為基準模型（benchmark），後續所有模型都與它比較。

**步驟**：

1. 請 AI 將資料分成 80% 訓練集和 20% 測試集（設定 `random_state=42` 確保可重現）
2. 類別變數做 one-hot encoding，處理缺失值
3. 用所有特徵建立 Logistic Regression 模型，預測 `OUTCOME`
4. 記錄模型的 Accuracy、Precision、Recall、F1-Score

**Prompt 參考**：
> 「請用 `Car_Insurance_Claim.csv` 建立一個 Logistic Regression 分類模型，預測 `OUTCOME`。步驟：(1) 處理缺失值（數值欄位用中位數填補）(2) 類別變數做 one-hot encoding (3) 將資料 80/20 分成訓練集和測試集（random_state=42）(4) 訓練模型並用測試集預測 (5) 顯示各變數的係數和 p-value (6) 印出 Accuracy、Precision、Recall、F1-Score (7) 畫出混淆矩陣。」

**思考題**：
- 哪些變數的係數最大？方向是否符合直覺？
- Accuracy 是多少？記下來，作為後續模型比較的基準
- 混淆矩陣中，False Negative 和 False Positive 各有多少？

---

### 任務 5：認識其他分類模型（講師口述介紹）

> 此段由講師口頭講解，學生聆聽並理解各模型的核心概念。

#### Decision Tree（決策樹）

- **核心概念**：像一棵倒過來的樹，每個節點問一個 Yes/No 問題，逐步將資料分群
- **優點**：結果直觀易解釋，可以畫出決策流程圖
- **缺點**：容易過擬合（overfitting），對資料小變動敏感
- **產險應用**：例如「年齡 < 25？→ 有超速紀錄？→ 高風險」

#### Random Forest（隨機森林）

- **核心概念**：建立很多棵決策樹，每棵樹用不同的資料子集，最後投票決定結果
- **優點**：比單棵決策樹穩定，不容易過擬合，可計算特徵重要性
- **缺點**：不像決策樹那麼容易解釋
- **產險應用**：找出哪些特徵對理賠預測最重要（特徵重要性排序）

#### SVM（支持向量機）

- **核心概念**：在資料點之間找一條「最佳分隔線」，把兩類資料分開，且間距最大化
- **優點**：在高維度資料表現好，可用不同核函數（linear、rbf）處理非線性問題
- **缺點**：資料量大時訓練較慢，結果較難解釋
- **產險應用**：適合特徵多但樣本相對少的情境

#### XGBoost（極端梯度提升）

- **核心概念**：一棵一棵樹依序建立，每棵新樹專門修正前一棵樹的錯誤，逐步提升準確度
- **優點**：目前業界最熱門的表格資料模型之一，速度快、精度高
- **缺點**：需要調整較多參數（如學習率、樹的深度）
- **產險應用**：Kaggle 競賽與業界風險建模的主流選擇

#### 模型比較重點

| 模型 | 可解釋性 | 準確度潛力 | 訓練速度 | 過擬合風險 |
|------|----------|------------|----------|------------|
| Logistic Regression | 高 | 中 | 快 | 低 |
| Decision Tree | 高 | 中 | 快 | 高 |
| Random Forest | 中 | 高 | 中 | 低 |
| SVM | 低 | 中高 | 慢 | 中 |
| XGBoost | 中 | 高 | 中 | 低 |

---

### 任務 6：Vibe Coding 實作 — 自選模型挑戰

**目標**：用 Vibe Coding 自行實作至少兩個分類模型，與 benchmark（Logistic Regression）比較。

**步驟**：

1. 從 Decision Tree、Random Forest、SVM、XGBoost 中**自選至少兩個模型**
2. 用自然語言 Prompt 請 AI 建立模型
3. 使用與任務 4 **相同的訓練/測試集切分**（random_state=42），確保公平比較
4. 記錄每個模型的 Accuracy、Precision、Recall、F1-Score
5. 將所有模型的結果整理成比較表

**Prompt 參考（依模型替換）**：

> Decision Tree：
> 「請用與之前相同的訓練集和測試集，建立一個 Decision Tree 分類模型預測 `OUTCOME`。請畫出決策樹圖（限制深度為 5 以便閱讀），並印出 Accuracy、Precision、Recall、F1-Score 和混淆矩陣。」

> Random Forest：
> 「請用與之前相同的訓練集和測試集，建立一個 Random Forest 分類模型（100 棵樹）預測 `OUTCOME`。請畫出特徵重要性排序圖（前 10 名），並印出 Accuracy、Precision、Recall、F1-Score 和混淆矩陣。」

> SVM：
> 「請用與之前相同的訓練集和測試集，建立一個 SVM 分類模型（使用 rbf kernel）預測 `OUTCOME`。請先對特徵做標準化（StandardScaler），並印出 Accuracy、Precision、Recall、F1-Score 和混淆矩陣。」

> XGBoost：
> 「請用與之前相同的訓練集和測試集，建立一個 XGBoost 分類模型預測 `OUTCOME`。請畫出特徵重要性排序圖（前 10 名），並印出 Accuracy、Precision、Recall、F1-Score 和混淆矩陣。」

**最後，請 AI 統整比較**：
> 「請將 Logistic Regression、[你選的模型 1]、[你選的模型 2] 的 Accuracy、Precision、Recall、F1-Score 整理成一張比較表，並畫成長條圖。哪個模型表現最好？」

**思考題**：
- 哪個模型的 Accuracy 最高？哪個模型的 Recall 最高？
- 最高 Accuracy 的模型一定是最好的嗎？為什麼？
- Decision Tree 的決策樹圖中，第一個分裂的變數是什麼？這代表它認為哪個特徵最重要
- Random Forest 或 XGBoost 的特徵重要性排序，跟 Logistic Regression 的係數大小排序一樣嗎？
- 如果你是保險公司主管，你會選哪個模型？考量點是什麼（準確度？可解釋性？）

---

### 任務 7：模型結果的商業解讀

**目標**：選擇最適合的模型，將分析結果轉化為商業建議，產出精算洞察報告。

**步驟**：

1. 從所有模型中選擇一個你認為最適合產險應用的模型，並說明理由
2. 從該模型的結果中，挑選 2-3 個最具鑑別度的風險特徵
3. 請 AI 協助生成一份 One-Pager 報告

**Prompt 參考**：
> 「根據前面所有模型的分析結果，請幫我生成一份精算洞察報告。格式如下：(1) 數據摘要：資料集基本資訊與使用的模型 (2) 模型選擇：為什麼選這個模型（比較表佐證）(3) 關鍵發現：列出 3 個最重要的風險因子及其影響 (4) 高風險客群：描述具體的特徵組合 (5) 保費調整建議：具體的調整方向與幅度。」

**注意事項**：
- 報告中的數字必須來自你的分析結果，不能憑空捏造
- 思考「公平待客原則」：哪些變數可以用來定價？哪些不行（例如種族）？
- 模型選擇要考量可解釋性 — 監理機關可能要求你解釋為什麼拒保或加費
- 建議是否具有實務可行性？

---

## 評分標準

| 項目 | 配分 | 說明 |
|------|------|------|
| 任務完成度 | 40% | 是否完成所有任務步驟 |
| 驗證與思考 | 30% | 是否回答思考題、是否驗證 AI 輸出 |
| 商業洞察 | 20% | 報告的分析深度與建議合理性 |
| Prompt 品質 | 10% | Prompt 是否清晰、有效率 |

---

## 提醒

1. **AI 是工具，不是答案**：AI 產出的結果一定要驗證，不能直接照抄
2. **統計素養是根本**：你需要判斷 AI 的統計解讀是否正確
3. **問好問題比寫程式更重要**：Prompt 的品質決定了分析的品質
4. **存檔**：請定期將 Antigravity 的程式碼和圖表截圖存檔，作為繳交作業的依據

# E-commerce Dashboard 可視化儀表板

## 簡介

這是一個基於 Streamlit 的交互式電商數據可視化儀表板，用於分析 2011年1月至11月的電商數據。儀表板提供全面的業務指標分析、客戶細分和退貨分析功能。

## 功能特點

### 1. KPI 概覽卡片（第一區塊）

儀表板頂部顯示關鍵業務指標，分為三行：

**第一行：核心業務指標**
- **Revenue**（總收入）- 顯示最後一個月的收入，包含 MoM（月環比）增長率
- **Orders**（訂單數）- 最後一個月的正常訂單數
- **Customers**（客戶數）- 最後一個月的客戶數

**第二行：平均指標**
- **AOV**（平均訂單價值）- Average Order Value
- **ARPU**（平均每用戶收入）- Average Revenue Per User

**第三行：退貨指標**
- **Return Amount**（退貨金額）- 最後一個月的退貨金額
- **Return Orders**（退貨訂單數）- 最後一個月的退貨訂單數
- **Return Rate**（退貨率）- 退貨訂單佔總訂單的比例

### 2. 月度趨勢圖表 (MOM)（第二區塊）

**Revenue & Orders 趨勢**
- 雙Y軸線圖，同時顯示收入和訂單數的月度變化
- 標記負增長月份（紅色X標記）

**Customers 趨勢**
- 柱狀圖顯示客戶數的月度變化

**AOV & ARPU 趨勢**
- 左右並排顯示兩個獨立的線圖
- 左側：AOV 趨勢
- 右側：ARPU 趨勢

### 3. RFM 客戶細分可視化（第三區塊）

**GUEST vs Others 比較**
- Monetary 長條圖：比較 GUEST 客戶和註冊客戶的總收入
- Count 長條圖：比較 GUEST 客戶和註冊客戶的數量
- 提示：已排除 GUEST 客戶，後續分析僅包含註冊客戶

**RFM 散點圖**
- **Total Score vs Revenue** 散點圖
- 顏色根據 RFM 類別區分（從 Champions 到 Lost：深藍到深紅）
- 顏色映射：
  - Champions: 深藍色 (#1a237e)
  - Loyal: 藍色 (#3949ab)
  - Potential Loyalist: 淺藍色 (#5c6bc0)
  - At Risk: 橙色 (#e64a19)
  - Lost: 深紅色 (#c62828)
  - Unknown: 灰色 (#95a5a6)

**Revenue Contribution**
- 餅圖顯示各 RFM 類別的收入貢獻
- 顯示詳細占比（金額和百分比）
- 顏色從 Champions 到 Lost：深藍到深紅

**Customer Contribution**
- 餅圖顯示各 RFM 類別的客戶數量貢獻
- 顯示詳細占比（人數和百分比）
- 顏色從 Champions 到 Lost：深藍到深紅

### 4. 退貨分析（第四區塊）

**Return Rate & Return Amount 趨勢**
- 雙Y軸圖表，同一張圖顯示兩個指標
- **Return Amount**：柱狀圖（主Y軸，紅色）
- **Return Rate**：線圖（次Y軸，藍色）
- 數據來源：MOM 月度數據

**Product Return Analysis**
- 散點圖：Return Amount vs Return Rate
- 顏色根據退貨類別區分（High/Medium/Low/Outlier）
- 氣泡大小表示 Return Count
- 顯示 StockCode 和退貨次數

**Customer Return Analysis**
- 散點圖：Return Amount vs Return Rate
- 顏色根據退貨類別區分（High/Medium/Low/Outlier）
- 氣泡大小表示 Return Count
- 顯示 CustomerID 和退貨次數

### 5. 自動生成可執行洞察
- 異常退貨高峰月份識別
- 客戶流失風險分析
- 高損失產品識別

## 安裝步驟

### 1. 安裝依賴

```bash
pip install -r requirements_visualization.txt
```

或者手動安裝：

```bash
pip install streamlit pandas numpy plotly openpyxl
```

### 2. 準備數據文件

確保以下文件存在於同一目錄：

- `彙總表.xlsx` - 包含以下工作表：
  - `MOM` - 月度 KPI 數據
  - `AOV_ARPU` - AOV 和 ARPU 數據
  - `RFM` - RFM 分析數據
  - `SKU` - SKU 數據
  - `Sales by Country` - 國家銷售數據

- `Return and Abnormal_2011_11.xlsx` (可選) - 包含：
  - `Return analysis product` - 產品退貨分析
  - `Abnormal analysis product` - 異常產品分析

### 3. 運行儀表板

```bash
streamlit run visualization_dashboard.py
```

儀表板將在瀏覽器中自動打開，通常地址為：`http://localhost:8501`

## 數據篩選

儀表板自動篩選 2011年1月至11月的數據（排除12月），確保分析的時間範圍一致。

## 使用說明

### 啟動儀表板

使用 `run_dashboard.py` 腳本啟動：

```bash
python run_dashboard.py
```

或直接使用 Streamlit：

```bash
streamlit run visualization_dashboard.py
```

### 功能導覽

1. **查看 KPI 概覽**（第一區塊）
   - 頁面頂部顯示關鍵指標卡片
   - 第一行：核心業務指標（Revenue, Orders, Customers）
   - 第二行：平均指標（AOV, ARPU）
   - 第三行：退貨指標（Return Amount, Return Orders, Return Rate）

2. **分析月度趨勢**（第二區塊）
   - Revenue & Orders 趨勢：雙Y軸線圖，查看收入和訂單的月度變化
   - Customers 趨勢：柱狀圖顯示客戶數變化
   - AOV & ARPU 趨勢：左右並排，對比平均訂單價值和每用戶收入

3. **客戶細分分析**（第三區塊）
   - GUEST vs Others：比較訪客客戶和註冊客戶
   - RFM 散點圖：Total Score vs Revenue，了解不同客戶群體的價值
   - Revenue/Customer Contribution：餅圖查看各類別的收入和客戶占比

4. **退貨分析**（第四區塊）
   - Return Rate & Return Amount 趨勢：查看退貨率和退貨金額的月度變化
   - Product Return Analysis：識別高退貨率產品
   - Customer Return Analysis：識別高退貨率客戶

5. **查看洞察**：閱讀自動生成的可執行建議

## 數據文件要求

### 必需文件

**彙總表.xlsx** - 必須包含以下工作表：
- `MOM` - 月度 KPI 數據（包含 Revenue, Normal_Orders, Return_Orders, Return, Customer 等列）
- `AOV_ARPU` - AOV 和 ARPU 數據（包含 YearMonth, AOV, ARPU 等列）
- `RFM` - RFM 分析數據（包含 CustomerID, Recency, Frequency, Monetary, Total_Score, Category 等列）
- `SKU` - SKU 數據（可選）
- `Sales by Country` - 國家銷售數據（可選）

### 可選文件

**Return and Abnormal_2011_11.xlsx** (或 `Return and Abnormal.xlsx`) - 包含：
- `Return analysis product` - 產品退貨分析（包含 Return_Amount, Return_Rate, Return_Count, Category, StockCode 等列）
- `Return analysis customer` - 客戶退貨分析（包含 Return_Amount, Return_Rate, Return_Count, Category 等列）
- `Abnormal analysis product` - 異常產品分析（可選）

## 注意事項

- 確保已運行 `execute_prompt.py` 生成 `彙總表.xlsx`
- 如需查看退貨分析，請確保有 `Return and Abnormal_2011_11.xlsx` 文件
- 如果文件不存在，相關部分會顯示警告信息
- Dashboard 自動篩選 2011年1月至11月的數據（排除12月）
- GUEST 客戶在 RFM 分析中被排除，但會單獨顯示比較

## 技術棧

- **Streamlit** - Web 應用框架
- **Plotly** - 交互式圖表庫
  - `plotly.express` - 快速創建圖表
  - `plotly.graph_objects` - 高級圖表定制
  - `plotly.subplots` - 子圖和雙Y軸支持
- **Pandas** - 數據處理
- **NumPy** - 數值計算
- **openpyxl** - Excel 文件讀取

## 主要函數說明

### `load_data()`
- 加載所有必需的數據文件
- 支持多種文件名格式（自動嘗試不同文件名）
- 返回包含所有數據的字典

### `filter_2011_data(df, date_column='YearMonth')`
- 篩選 2011年1月至11月的數據
- 排除 2011年12月
- 支持不同的日期格式

### `generate_kpi(data)`
- 生成 KPI 概覽卡片
- 顯示最後一個月的數據
- 計算 MoM 增長率

### `generate_mom_charts(data)`
- 生成月度趨勢圖表
- 包括 Revenue & Orders、Customers、AOV、ARPU

### `generate_rfm_visualization(data)`
- 生成 RFM 客戶細分可視化
- 包括 GUEST vs Others 比較、RFM 散點圖、餅圖

### `generate_return_analysis(data)`
- 生成退貨分析可視化
- 包括趨勢圖和散點圖

### `generate_insights(data)`
- 自動生成可執行洞察
- 識別異常退貨、客戶流失風險、高損失產品

## 圖表說明

### KPI 卡片
- **Revenue MoM**：顯示月環比增長率（相對於前一個月）
- **Return Rate**：計算公式 = Return_Orders / (Return_Orders + Normal_Orders) × 100%
- **AOV**：計算公式 = Revenue / Normal_Orders
- **ARPU**：從 AOV_ARPU 數據讀取，或計算 = Revenue / Customers

### MOM 圖表
- **Revenue & Orders**：使用雙Y軸，左側Y軸顯示 Revenue（美元），右側Y軸顯示 Orders（訂單數）
- **Customers**：柱狀圖，綠色顯示
- **AOV & ARPU**：分開顯示，便於對比

### RFM 可視化
- **GUEST 識別**：CustomerID 為 "GUEST"（不區分大小寫）的客戶被識別為訪客客戶
- **RFM 散點圖**：X軸為 Total Score，Y軸為 Revenue（Monetary），顏色根據 Category 區分
- **餅圖**：顯示各類別的收入和客戶占比，顏色從 Champions（深藍）到 Lost（深紅）

### Return Analysis
- **Return Rate & Return Amount**：雙Y軸圖表，左側Y軸顯示 Return Amount（美元），右側Y軸顯示 Return Rate（百分比）
- **散點圖**：X軸為 Return Amount，Y軸為 Return Rate，氣泡大小表示 Return Count

## 自定義

可以根據需要修改：
- 時間範圍篩選（修改 `filter_2011_data` 函數）
- 圖表樣式和顏色（修改各圖表的 `color_discrete_map` 或 `marker` 參數）
- KPI 計算邏輯（修改 `generate_kpi` 函數）
- 洞察生成規則（修改 `generate_insights` 函數）
- RFM 顏色映射（修改 `color_map` 字典）

## 故障排除

### 問題：無法加載數據
- 檢查文件路徑是否正確
- 確認文件存在且格式正確
- 查看終端錯誤信息

### 問題：圖表不顯示
- 確認數據文件包含必要的列
- 檢查數據格式是否正確
- 查看瀏覽器控制台錯誤

### 問題：Streamlit 無法啟動
- 確認已安裝所有依賴
- 檢查 Python 版本（建議 3.10 或 3.11）
- 嘗試重新安裝 streamlit
- 使用 `run_dashboard.py` 腳本啟動

### 問題：RFM 圖表不顯示
- 確認 RFM 數據包含必要的列（CustomerID, Recency, Frequency, Monetary, Total_Score, Category）
- 檢查是否有 GUEST 客戶（會被排除）
- 確認數據格式正確

### 問題：Return Analysis 不顯示
- 確認有 `Return and Abnormal_2011_11.xlsx` 文件
- 檢查文件是否包含 `Return analysis product` 和 `Return analysis customer` 工作表
- 確認 MOM 數據包含 Return_Orders 和 Return 列

### 問題：KPI 卡片顯示為 0
- 確認 MOM 數據包含最後一個月的數據
- 檢查 YearMonth 格式是否正確
- 確認數據已正確篩選（2011年1-11月）






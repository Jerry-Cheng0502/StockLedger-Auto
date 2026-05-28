# 國泰證券自動化對帳單解析與 FIFO 損益記帳系統

一個基於 Python 的自動化理財工具。本專案透過 **`uv`** 進行執行環境管理，能自動登入 Gmail 批次下載指定日期範圍的「國泰證券日對帳單」PDF，解密並解析內部交易表格，再利用**先進先出 (FIFO) 算法**自動沖銷庫存，精確計算出每筆交易的買進/賣出時間與實現損益，同時追蹤未賣出的庫存成本，最後匯出成多工作表的 Excel 報表。

## 🌟 功能特色

- **Gmail 批次下載**：依郵件寄送日期範圍，一次下載所有符合條件的國泰證券 PDF 對帳單，已下載的自動略過，不重複抓取。
- **批次 PDF 解析**：支援指定單一 PDF 或掃描整個資料夾，一次處理多個對帳單。
- **安全 PDF 解密**：支援輸入身分證字號自動解密受密碼保護的國泰對帳單。
- **智慧防呆機制**：自動比對 Excel 歷史紀錄，若對帳單日期已存在則略過，避免重複寫入與計算。
- **核心 FIFO 損益模型**：自動配對買賣股票，精確拆解並追蹤「買進時間」與「賣出時間」，計算包含手續費及交易稅的精確「總獲利」與「報酬率」。
- **未實現庫存追蹤**：自動彙整尚未賣出的持股，列出買進時間、持有股數與帳面成本。
- **多維度 Excel 報表**：
  - `實現損益總表`：FIFO 配對損益，依賣出時間排序。
  - `未實現庫存`：現有持股成本，依買進時間排序。
  - `交易明細歷史`：完整的原始交易底稿。
  - `月報` / `年報`：自動彙整各期間損益、交易次數與費用。
  - `個股彙整`：各股票累計損益與平均報酬率排行。
- **終端機摘要**：執行後直接在 terminal 顯示損益概況、前三名個股，免開 Excel 也能一目瞭然。

---

## 🛠️ 開發環境與套件

腳本採用 `uv` inline script metadata，**不需要建立專案或 `pyproject.toml`**，直接 `uv run` 即可自動安裝所有依賴：

- **`pdfplumber`**：精確解析 PDF 表格與文字內容。
- **`pandas`**：核心數據清洗與 FIFO 計算。
- **`openpyxl`**：Excel 活頁簿寫入與欄寬調整。
- **`python-dotenv`**：環境變數與敏感憑證管理。

---

## 🚀 快速開始

### 1. 安裝 `uv`

若尚未安裝 `uv`，請參考 [官方文件](https://docs.astral.sh/uv/getting-started/installation/) 安裝：

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### 2. 環境設定 (`.env`)

在腳本同目錄下建立 `.env` 檔案，填入個人憑證：

```env
# 國泰對帳單 PDF 解密密碼（通常為身分證字號，英文字母需大寫）
CATHAY_ID=A123456789

# Gmail 自動下載設定（選填，若不填則手動將 PDF 放至目錄）
GMAIL_USER=your_email@gmail.com
# 請填入 Gmail 的「應用程式密碼」，而非一般登入密碼
GMAIL_APP_PASSWORD=abcd efgh ijkl mnop
# 可選：自訂寄件方（預設為國泰官方寄件地址）
# GMAIL_SENDER=e-notification@ebill1.cathaysec.com.tw
```

### 3. 執行程式

`uv run` 會自動讀取腳本頂部的依賴宣告，首次執行時自動安裝套件，之後快取沿用：

```bash
uv run main.py
```

---

## 📖 使用方式

### 輸入來源（三選一）

```bash
# 指定單一 PDF
uv run main.py --pdf 國泰證券日對帳單.pdf

# 批次掃描資料夾內所有 PDF
uv run main.py --folder ./pdfs/

# 從 Gmail 下載（見下方說明）
uv run main.py --gmail
```

### Gmail 批次下載

```bash
# 下載所有歷史郵件
uv run main.py --gmail

# 只下載特定日期範圍的郵件
uv run main.py --gmail --gmail-start 2024/01/01 --gmail-end 2024/12/31

# 指定 PDF 存放資料夾（預設：./gmail_pdfs）
uv run main.py --gmail --gmail-start 2025/01/01 --gmail-save-dir ./2025_pdfs
```

> `--gmail-start` / `--gmail-end` 是依據**郵件寄送日期**篩選，與損益報表的時間範圍篩選（`--start` / `--end`）是獨立的兩組參數。

### 損益篩選與輸出控制

```bash
# 只看特定時間範圍的損益（不影響 FIFO 計算邏輯，只影響報表顯示）
uv run main.py --start 2024/01/01 --end 2024/12/31

# 只看特定股票（支援部分比對）
uv run main.py --stock 台積電

# 指定 Excel 輸出路徑
uv run main.py --output ./reports/2024.xlsx

# Dry-run：只印出摘要，不寫入 Excel
uv run main.py --dry-run
```

### 完整參數列表

| 參數 | 說明 |
|------|------|
| `--pdf FILE` | 指定單一 PDF 路徑 |
| `--folder DIR` | 批次掃描資料夾內所有 PDF |
| `--gmail` | 從 Gmail 批次下載對帳單 |
| `--gmail-start YYYY/MM/DD` | Gmail 篩選起始日（含） |
| `--gmail-end YYYY/MM/DD` | Gmail 篩選結束日（含） |
| `--gmail-save-dir DIR` | PDF 下載存放資料夾（預設：`./gmail_pdfs`） |
| `--start YYYY/MM/DD` | 損益報表篩選起始日（含） |
| `--end YYYY/MM/DD` | 損益報表篩選結束日（含） |
| `--stock NAME` | 只顯示特定股票（部分比對） |
| `--output FILE` | Excel 輸出路徑（預設：`股票對帳單總表.xlsx`） |
| `--password ID` | 身分證字號（不傳則從 `.env` 或互動輸入） |
| `--dry-run` | 只印摘要，不寫入 Excel |

---

## 📊 資料流與 FIFO 模型說明

```
Gmail / 本地 PDF
       │
       ▼
  PDF 解密與解析
  （提取交易日期與明細表格）
       │
       ▼
  防呆檢查
  （跳過已寫入過的日期）
       │
       ▼
  合併歷史明細
  （新資料 + 交易明細歷史）
       │
       ▼
  FIFO 沖銷引擎
  ┌────────────────────────┐
  │ 買進 → 推入庫存池      │
  │ 賣出 → 由舊到新配對    │
  │        計算實現損益    │
  │ 剩餘 → 未實現庫存      │
  └────────────────────────┘
       │
       ▼
  彙整報表（月報/年報/個股）
       │
       ▼
  寫入 Excel（多工作表）
```

---

## 📝 注意事項

1. **Gmail 應用程式密碼**：需至 [Google 帳戶安全性設定](https://myaccount.google.com/apppasswords) 申請專用的應用程式密碼，並確保帳戶已開啟 IMAP 存取功能。一般登入密碼無法使用。
2. **手動模式**：若不使用 Gmail，只需將對帳單 PDF 放至指定路徑或資料夾，以 `--pdf` 或 `--folder` 參數執行即可。
3. **檔案佔用**：執行前請確保 Excel 檔案處於關閉狀態，避免 Windows 檔案鎖定導致寫入失敗。
4. **損益篩選不影響計算**：`--start` / `--end` 只篩選輸出報表的顯示範圍，FIFO 計算仍基於完整歷史資料，確保庫存配對的正確性。

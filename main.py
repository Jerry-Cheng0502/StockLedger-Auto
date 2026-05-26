import os
import re
import pdfplumber
import imaplib
import email
from email.header import decode_header
import pandas as pd
import numpy as np
from dotenv import load_dotenv

# 嘗試自動載入 .env
try:
    load_dotenv()
except ImportError:
    pass

ID_PASSWORD = os.getenv("CATHAY_ID")
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")

if not ID_PASSWORD:
    ID_PASSWORD = input("請輸入你的身分證字號（字母需大寫）：").strip()

PDF_INPUT_PATH = "國泰證券日對帳單.pdf"
EXCEL_OUTPUT_PATH = "股票對帳單總表.xlsx"

# =====================================================================
# 📨 0. 自動從 Gmail 下載最新對帳單 PDF
# =====================================================================
def fetch_latest_pdf_from_gmail():
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        print("⚠️ 未偵測到 Gmail 設定，跳過自動下載，直接讀取本地檔案。")
        return False

    print("📬 正在連線至 Gmail 尋找國泰對帳單...")
    try:
        # 連線至 Gmail IMAP 伺服器
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        mail.select("inbox")

        # 搜尋郵件：來自國泰證券(一般為 cathaysec) 且 標題包含「日對帳單」
        # 你可以根據實際收到的信件主旨調整搜尋關鍵字
        search_criterion = '(FROM "e-notification@ebill1.cathaysec.com.tw")'
        status, messages = mail.search('UTF-8', search_criterion)
        
        if status != "OK" or not messages[0]:
            print("⚠️ 未在 Gmail 中找到符合的國泰日對帳單郵件。")
            mail.logout()
            return False

        # 取得所有符合條件的郵件 ID，並拿最後一封（最新的）
        mail_ids = messages[0].split()
        latest_mail_id = mail_ids[-1]

        # 抓取該封郵件內容
        status, data = mail.fetch(latest_mail_id, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        # 解析郵件附件
        pdf_downloaded = False
        for part in msg.walk():
            if part.get_content_maintype() == "multipart" or part.get("Content-Disposition") is None:
                continue

            filename = part.get_filename()
            if filename:
                decode_res = decode_header(filename)[0]
                if isinstance(decode_res[0], bytes):
                    filename = decode_res[0].decode(decode_res[1] or "utf-8")
                
                if filename.lower().endswith(".pdf"):
                    print(f"📥 發現最新對帳單信件，正在下載附件: {filename}...")
                    with open(PDF_INPUT_PATH, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    pdf_downloaded = True
                    break 

        mail.logout()
        return pdf_downloaded

    except Exception as e:
        print(f"❌ Gmail 檔案抓取失敗: {str(e)}")
        return False

# 執行自動下載
download_success = fetch_latest_pdf_from_gmail()
if not download_success and not os.path.exists(PDF_INPUT_PATH):
    print("❌ 既無法從 Gmail 下載，本地也沒有發現對帳單檔案，程式結束。")
    exit()

# =====================================================================
# 📂 預先讀取歷史紀錄（供防呆檢查使用）
# =====================================================================
history_dates = set()
if os.path.exists(EXCEL_OUTPUT_PATH):
    try:
        df_history = pd.read_excel(EXCEL_OUTPUT_PATH, sheet_name="交易明細歷史")
        if not df_history.empty and "時間" in df_history.columns:
            # 轉為字串格式存入 set 方便比對
            history_dates = set(df_history["時間"].astype(str).unique())
    except Exception as e:
        print(f"⚠️ 讀取歷史 Excel 失敗（可能檔案被打開或損壞）: {e}")
        df_history = pd.DataFrame()
else:
    df_history = pd.DataFrame()

# =====================================================================
# 📊 1. 從 PDF 讀取當日原始明細
# =====================================================================
print("🔓 正在解析對帳單...")
raw_rows = []

try:
    with pdfplumber.open(PDF_INPUT_PATH, password=ID_PASSWORD) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            text_clean = re.sub(r'\s+', '', text)
            date_match = re.search(r'證券日對帳單(\d{2,3})年(\d{1,2})月(\d{1,2})日', text_clean)
            if date_match:
                roc_year, month, day = date_match.groups()
                ad_year = int(roc_year) + 1911
                trade_date = f"{ad_year}/{int(month):02d}/{int(day):02d}"
                print(f"✅ 對帳單日期：{trade_date}")
                
                if trade_date in history_dates:
                    print(f"🛑 [防呆提示] 最新下載的對帳單日期為 {trade_date}。")
                    print(f"   此日期的交易紀錄先前「已經成功寫入」Excel。")
                    print("   為了防止重複計算，程式已安全終止，未對 Excel 進行任何變更。")
                    exit()
                
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and len(row) >= 7:
                        prod_name = str(row[0]).strip()
                        category = str(row[1]).strip()
                        if "商品名稱" in prod_name or "總合計" in prod_name or "成交日期" in prod_name or not row[0]:
                            continue
                        if "買" in category or "賣" in category:
                            raw_rows.append(row)
except Exception as e:
    print(f"❌ PDF 處理或防呆檢查失敗: {str(e)}")
    exit()

if not raw_rows:
    print("⚠️ 未發現有效交易明細。")
    exit()

# 整理成當日新交易的 DataFrame
new_details = []
for row in raw_rows:
    action = "買進" if "買" in str(row[1]) else "賣出"
    new_details.append({
        "時間": trade_date,
        "商品名稱": row[0].strip(),
        "買賣別": action,
        "成交股數": int(str(row[2]).replace(",", "")),
        "單價": float(str(row[3]).replace(",", "")),
        "手續費": int(str(row[5]).replace(",", "")),
        "交易稅": int(str(row[6]).replace(",", ""))
    })
df_new = pd.DataFrame(new_details)

# =====================================================================
# 📂 2. 讀取歷史總庫存紀錄 (確保歷史連續性)
# =====================================================================
if os.path.exists(EXCEL_OUTPUT_PATH):
    try:
        df_history = pd.read_excel(EXCEL_OUTPUT_PATH, sheet_name="交易明細歷史")
        df_all_details = pd.concat([df_history, df_new], ignore_index=True).drop_duplicates()
    except Exception:
        df_all_details = df_new
else:
    df_all_details = df_new

# =====================================================================
# 🧠 3. 核心算法：先進先出 (FIFO) 股票損益沖銷模型 (升級時間追蹤)
# =====================================================================
print("🧠 正在重新進行股票買賣配對與損益計算...")
realized_profit_loss = []

for stock_name, group in df_all_details.groupby("商品名稱"):
    buy_pool = []
    # 確保按時間由舊到新排序處理
    for _, row in group.sort_values(by="時間").iterrows():
        if row["買賣別"] == "買進":
            buy_pool.append({
                "date": row["時間"],  # 這裡精確記錄了買進時間
                "shares": row["成交股數"],
                "price": row["單價"],
                "fee_per_share": row["手續費"] / row["成交股數"]
            })
        elif row["買賣別"] == "賣出":
            sell_shares = row["成交股數"]
            sell_price = row["單價"]
            sell_fee_total = row["手續費"]
            sell_tax_total = row["交易稅"]
            sell_date = row["時間"]  # 當前賣出交易的時間
            
            while sell_shares > 0 and buy_pool:
                current_buy = buy_pool[0]
                matched_shares = min(sell_shares, current_buy["shares"])
                
                match_buy_fee = round(current_buy["fee_per_share"] * matched_shares)
                match_sell_fee = round((sell_fee_total / row["成交股數"]) * matched_shares)
                match_sell_tax = round((sell_tax_total / row["成交股數"]) * matched_shares)
                
                buy_cost = round(current_buy["price"] * matched_shares) + match_buy_fee
                sell_revenue = round(sell_price * matched_shares) - match_sell_fee - match_sell_tax
                
                total_profit = sell_revenue - buy_cost
                return_rate = (total_profit / buy_cost) if buy_cost > 0 else 0
                
                # ✨【變更重點】將單一時間欄位，拆解為「買進時間」與「賣出時間」
                realized_profit_loss.append({
                    "商品名稱": stock_name,
                    "股數": matched_shares,
                    "買進時間": current_buy["date"],  # 來自庫存池的日期
                    "賣出時間": sell_date,           # 來自本次賣出的日期
                    "買進價格": current_buy["price"],
                    "賣出價格": sell_price,
                    "手續費(買+賣)": match_buy_fee + match_sell_fee,
                    "交易稅": match_sell_tax,
                    "總成本": buy_cost,
                    "總獲利": total_profit,
                    "報酬率": f"{return_rate*100:.2f}%"
                })
                
                sell_shares -= matched_shares
                current_buy["shares"] -= matched_shares
                if current_buy["shares"] == 0:
                    buy_pool.pop(0)

df_profit_loss = pd.DataFrame(realized_profit_loss)

# =====================================================================
# 💾 4. 寫入 Excel (雙工作表：一份留歷史底稿，一份看精美損益)
# =====================================================================
with pd.ExcelWriter(EXCEL_OUTPUT_PATH, engine="openpyxl") as writer:
    df_profit_loss.to_excel(writer, sheet_name="實現損益總表", index=False)
    df_all_details.to_excel(writer, sheet_name="交易明細歷史", index=False)

print(f"🎉 終極對帳單更新成功！請查看 Excel 中的【實現損益總表】工作表！")
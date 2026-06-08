"""
國泰證券對帳單解析器（升級版）
===============================
功能：
  - 支援 CLI 參數（日期範圍、模式、路徑等）
  - 批次處理資料夾內所有 PDF
  - 計算「實現損益」與「未實現損益（現有庫存）」
  - 月報 / 年報自動彙整
  - Gmail 批次下載（依郵件寄送日期範圍，可選）

使用範例：
  # 基本用法（互動模式）
  uv run main.py

  # 指定單一 PDF
  uv run main.py --pdf 國泰證券日對帳單.pdf

  # 批次處理資料夾
  uv run main.py --folder ./pdfs/

  # 從 Gmail 下載「所有」國泰對帳單郵件
  uv run main.py --gmail

  # 從 Gmail 只下載特定日期範圍內的郵件（依郵件寄送日期）
  uv run main.py --gmail --gmail-start 2024/01/01 --gmail-end 2024/12/31

  # 只看特定損益範圍
  uv run main.py --start 2024/01/01 --end 2024/12/31

  # 不寫入 Excel，只印出摘要
  uv run main.py --dry-run
"""

import os
import re
import sys
import argparse
import imaplib
import email
from email.header import decode_header
from pathlib import Path
from datetime import datetime, date

import pdfplumber
import pandas as pd

# ── 嘗試載入 .env ──────────────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ══════════════════════════════════════════════════════════════════════════════
# ⚙️  設定常數（可透過環境變數覆寫）
# ══════════════════════════════════════════════════════════════════════════════
DEFAULT_PDF_PATH    = os.getenv("PDF_PATH",    "國泰證券日對帳單.pdf")
DEFAULT_PDF_FOLDER  = os.getenv("PDF_FOLDER",  "./pdfs")
DEFAULT_EXCEL_PATH  = os.getenv("EXCEL_PATH",  "股票對帳單總表.xlsx")
GMAIL_USER          = os.getenv("GMAIL_USER")
GMAIL_APP_PASSWORD  = os.getenv("GMAIL_APP_PASSWORD")
GMAIL_SENDER        = os.getenv("GMAIL_SENDER", "e-notification@ebill1.cathaysec.com.tw")

# ══════════════════════════════════════════════════════════════════════════════
# 🔧  CLI 參數解析
# ══════════════════════════════════════════════════════════════════════════════
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="國泰證券對帳單解析器",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    # 輸入來源
    src = parser.add_mutually_exclusive_group()
    src.add_argument("--pdf",    metavar="FILE",   help="指定單一 PDF 路徑")
    src.add_argument("--folder", metavar="DIR",    help="批次掃描資料夾內所有 PDF")
    src.add_argument("--gmail",  action="store_true", help="從 Gmail 批次下載國泰對帳單")

    # Gmail 郵件日期範圍（只在 --gmail 時有效）
    parser.add_argument("--gmail-start", metavar="YYYY/MM/DD",
                        help="只下載此日期之後寄出的郵件（依郵件寄送日，含）")
    parser.add_argument("--gmail-end",   metavar="YYYY/MM/DD",
                        help="只下載此日期之前寄出的郵件（依郵件寄送日，含）")
    parser.add_argument("--gmail-save-dir", metavar="DIR", default="./gmail_pdfs",
                        help="Gmail 下載的 PDF 存放資料夾（預設：./gmail_pdfs）")

    # 輸出
    parser.add_argument("--output", metavar="FILE", default=DEFAULT_EXCEL_PATH,
                        help=f"Excel 輸出路徑（預設：{DEFAULT_EXCEL_PATH}）")
    parser.add_argument("--dry-run", action="store_true",
                        help="只顯示摘要，不寫入 Excel")

    # 篩選
    parser.add_argument("--start", metavar="YYYY/MM/DD",
                        help="損益篩選起始日（含）")
    parser.add_argument("--end",   metavar="YYYY/MM/DD",
                        help="損益篩選結束日（含）")
    parser.add_argument("--stock", metavar="NAME",
                        help="只看特定股票（支援部分比對，例如：台積電）")

    # 報表模式
    parser.add_argument("--report", choices=["monthly", "yearly", "stock", "all"],
                        default="all", help="彙整報表類型（預設：all）")

    # 密碼
    parser.add_argument("--password", metavar="ID",
                        help="身分證字號（若不傳則從環境變數或互動輸入）")

    return parser.parse_args()


# ══════════════════════════════════════════════════════════════════════════════
# 📨  Gmail 批次下載（依郵件寄送日期範圍）
# ══════════════════════════════════════════════════════════════════════════════
def _imap_date(dt: datetime) -> str:
    """將 datetime 轉為 IMAP SINCE/BEFORE 格式，例如 '01-Jan-2024'"""
    return dt.strftime("%d-%b-%Y")


def fetch_pdfs_from_gmail(
    save_dir: str,
    start_date: str | None = None,
    end_date:   str | None = None,
) -> list[str]:
    """
    從 Gmail 下載所有符合條件的國泰對帳單 PDF。

    參數：
        save_dir   — PDF 儲存資料夾
        start_date — 郵件寄送起始日（"YYYY/MM/DD"），None 表示不限
        end_date   — 郵件寄送結束日（"YYYY/MM/DD"），None 表示不限

    回傳：下載成功的 PDF 路徑清單
    """
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        print("⚠️  未設定 GMAIL_USER / GMAIL_APP_PASSWORD，略過 Gmail 下載。")
        return []

    Path(save_dir).mkdir(parents=True, exist_ok=True)

    # ── 建立 IMAP 搜尋條件 ───────────────────────────────────────────────────
    criteria_parts = [f'FROM "{GMAIL_SENDER}"']

    if start_date:
        dt_start = datetime.strptime(start_date, "%Y/%m/%d")
        criteria_parts.append(f'SINCE "{_imap_date(dt_start)}"')
    if end_date:
        # IMAP BEFORE 是「嚴格小於」，所以要加一天
        from datetime import timedelta
        dt_end = datetime.strptime(end_date, "%Y/%m/%d") + timedelta(days=1)
        criteria_parts.append(f'BEFORE "{_imap_date(dt_end)}"')

    search_criterion = "(" + " ".join(criteria_parts) + ")"

    range_desc = []
    if start_date: range_desc.append(f"從 {start_date}")
    if end_date:   range_desc.append(f"到 {end_date}")
    range_str = "、".join(range_desc) if range_desc else "全部"
    print(f"📬 連線至 Gmail，搜尋國泰對帳單郵件（{range_str}）…")

    downloaded: list[str] = []

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        mail.select("inbox")

        status, messages = mail.search("UTF-8", search_criterion)
        if status != "OK" or not messages[0]:
            print("⚠️  未找到符合條件的郵件。")
            mail.logout()
            return []

        mail_ids = messages[0].split()
        print(f"  📧 共找到 {len(mail_ids)} 封符合郵件，開始逐一下載附件…")

        for idx, mail_id in enumerate(mail_ids, 1):
            try:
                _, data = mail.fetch(mail_id, "(RFC822)")
                msg = email.message_from_bytes(data[0][1])

                # 取得郵件寄送日期（供檔名使用）
                mail_date_str = msg.get("Date", "")
                try:
                    from email.utils import parsedate_to_datetime
                    mail_dt = parsedate_to_datetime(mail_date_str)
                    date_tag = mail_dt.strftime("%Y%m%d")
                except Exception:
                    date_tag = f"mail{idx:04d}"

                found_pdf = False
                for part in msg.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    if not part.get("Content-Disposition"):
                        continue

                    filename = part.get_filename()
                    if not filename:
                        continue

                    decoded, charset = decode_header(filename)[0]
                    if isinstance(decoded, bytes):
                        filename = decoded.decode(charset or "utf-8")

                    if not filename.lower().endswith(".pdf"):
                        continue

                    # 以「郵件日期_原始檔名」命名，避免衝突
                    safe_name = f"{date_tag}_{filename}"
                    save_path = os.path.join(save_dir, safe_name)

                    # 若已存在相同檔案則跳過（冪等）
                    if os.path.exists(save_path):
                        print(f"  [{idx}/{len(mail_ids)}] ⏭️  已存在，略過：{safe_name}")
                        downloaded.append(save_path)
                        found_pdf = True
                        break

                    with open(save_path, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    print(f"  [{idx}/{len(mail_ids)}] 📥 已下載：{safe_name}")
                    downloaded.append(save_path)
                    found_pdf = True
                    break

                if not found_pdf:
                    print(f"  [{idx}/{len(mail_ids)}] ⚠️  第 {idx} 封郵件無 PDF 附件，略過。")

            except Exception as e:
                print(f"  [{idx}/{len(mail_ids)}] ❌ 處理郵件失敗：{e}")
                continue

        mail.logout()
        print(f"\n✅ Gmail 下載完成，共取得 {len(downloaded)} 個 PDF，存放於：{save_dir}")
        return downloaded

    except Exception as e:
        print(f"❌ Gmail 連線失敗：{e}")
        return []


# ══════════════════════════════════════════════════════════════════════════════
# 📂  PDF 解析
# ══════════════════════════════════════════════════════════════════════════════
def parse_pdf(pdf_path: str, password: str) -> tuple[str | None, list[dict]]:
    """
    解析單一 PDF，回傳 (trade_date, raw_rows)。
    trade_date 格式為 "YYYY/MM/DD"；raw_rows 為原始表格列。
    """
    trade_date = None
    raw_rows: list[dict] = []

    try:
        with pdfplumber.open(pdf_path, password=password) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                text_clean = re.sub(r"\s+", "", text)

                # 解析民國日期
                if trade_date is None:
                    m = re.search(r"證券日對帳單(\d{2,3})年(\d{1,2})月(\d{1,2})日", text_clean)
                    if m:
                        ry, mo, dy = m.groups()
                        trade_date = f"{int(ry)+1911}/{int(mo):02d}/{int(dy):02d}"

                for table in page.extract_tables():
                    for row in table:
                        if not row or len(row) < 7:
                            continue
                        prod  = str(row[0]).strip()
                        categ = str(row[1]).strip()
                        if any(kw in prod for kw in ("商品名稱", "總合計", "成交日期")) or not row[0]:
                            continue
                        if "買" in categ or "賣" in categ:
                            raw_rows.append(row)

    except Exception as e:
        print(f"❌ 解析 {pdf_path} 失敗：{e}")
        return None, []

    return trade_date, raw_rows


def rows_to_df(trade_date: str, raw_rows: list) -> pd.DataFrame:
    records = []
    for row in raw_rows:
        # 國泰對帳單 row[1] 是類別（集買、集賣、沖買、沖賣...）
        categ = str(row[1]).strip()
        
        # 根據類別判斷它是買進還是賣出，供傳統 FIFO 判斷
        if "買" in categ:
            action = "買進"
        elif "賣" in categ:
            action = "賣出"
        else:
            continue  # 若有其他非買賣類別則跳過
            
        records.append({
            "時間":     trade_date,
            "商品名稱": row[0].strip(),
            "類別":     categ,  # 🌟 必須保留！這樣 run_fifo 才抓得到 "沖買"、"沖賣"
            "買賣別":   action, # 保留原有的買賣別
            "成交股數": int(str(row[2]).replace(",", "")),
            "單價":     float(str(row[3]).replace(",", "")),
            "手續費":   int(str(row[5]).replace(",", "")),
            "交易稅":   int(str(row[6]).replace(",", "")),
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════════════════
# 🧠  FIFO 損益引擎
# ══════════════════════════════════════════════════════════════════════════════
def run_fifo(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    對全部歷史明細跑 FIFO 配對。
    - 當沖（沖買/沖賣）：直接配對同一天的沖買與沖賣總量，合併計算損益，不進入庫存池。
    - 現股（集買/集賣）：集買進庫存池，集賣走 FIFO 消耗歷史庫存。
    回傳 (df_realized, df_unrealized)。
    """
    realized:  list[dict] = []
    inventory: list[dict] = []
 
    if df.empty:
        return pd.DataFrame(realized), pd.DataFrame(inventory)

    # 確保欄位存在（防禦性防護）
    if "類別" not in df.columns:
        # 萬一歷史 Excel 只有買賣別，建立相容欄位
        df["類別"] = df["買賣別"].replace({"買進": "集買", "賣出": "集賣"})

    for stock_name, group in df.groupby("商品名稱"):
        buy_pool: list[dict] = []
 
        for date, day in group.sort_values("時間").groupby("時間"):
 
            # ── 1. 當沖處理：沖買 + 沖賣 合併總計 ──────────────────────────────────
            intra_buys  = day[day["類別"] == "沖買"]
            intra_sells = day[day["類別"] == "沖賣"]
 
            if not intra_buys.empty and not intra_sells.empty:
                total_shares = intra_buys["成交股數"].sum()
                
                if total_shares > 0:
                    # 計算當沖加權平均單價
                    avg_buy_price  = (intra_buys["單價"] * intra_buys["成交股數"]).sum() / total_shares
                    avg_sell_price = (intra_sells["單價"] * intra_sells["成交股數"]).sum() / total_shares
                    
                    buy_fee  = intra_buys["手續費"].sum()
                    sell_fee = intra_sells["手續費"].sum()
                    sell_tax = intra_sells["交易稅"].sum()
                    
                    cost     = round(avg_buy_price * total_shares) + buy_fee
                    revenue  = round(avg_sell_price * total_shares) - sell_fee - sell_tax
                    profit   = revenue - cost
                    rr       = profit / cost if cost > 0 else 0
                    
                    realized.append({
                        "商品名稱":       stock_name,
                        "股數":          total_shares,
                        "買進時間":       date,
                        "賣出時間":       date,
                        "買進價格":       round(avg_buy_price, 2),
                        "賣出價格":       round(avg_sell_price, 2),
                        "手續費(買+賣)": buy_fee + sell_fee,
                        "交易稅":        sell_tax,
                        "總成本":        cost,
                        "總獲利":        profit,
                        "報酬率":        f"{rr*100:.2f}%",
                        "備註":          "當沖",
                    })

            # ── 2. 過濾掉當沖：讓剩下的「現股交易（集買/集賣）」去跑 FIFO 庫存計算 ──────
            # 這樣可以徹底避免當沖交易和現股交易混在一起算錯
            normal_day = day[~day["類別"].isin(["沖買", "沖賣"])]
 
            # ── 現股：現買進庫存池 ────────────────────────────────────────────
            for _, row in normal_day[normal_day["買賣別"] == "買進"].iterrows():
                if row["成交股數"] > 0:
                    buy_pool.append({
                        "date":          date,
                        "shares":        row["成交股數"],
                        "price":         row["單價"],
                        "fee_per_share": row["手續費"] / row["成交股數"],
                    })
 
            # ── 現股：現賣走 FIFO ─────────────────────────────────────────────
            for _, row in normal_day[normal_day["買賣別"] == "賣出"].iterrows():
                sell_shares  = row["成交股數"]
                sell_price   = row["單價"]
                sell_fee_tot = row["手續費"]
                sell_tax_tot = row["交易稅"]
 
                while sell_shares > 0 and buy_pool:
                    cur   = buy_pool[0]
                    msh   = min(sell_shares, cur["shares"])
                    ratio = msh / row["成交股數"]
 
                    buy_fee  = round(cur["fee_per_share"] * msh)
                    sell_fee = round(sell_fee_tot * ratio)
                    sell_tax = round(sell_tax_tot * ratio)
 
                    cost    = round(cur["price"] * msh) + buy_fee
                    revenue = round(sell_price   * msh) - sell_fee - sell_tax
                    profit  = revenue - cost
                    rr      = profit / cost if cost > 0 else 0
 
                    realized.append({
                        "商品名稱":       stock_name,
                        "股數":          msh,
                        "買進時間":       cur["date"],
                        "賣出時間":       date,
                        "買進價格":       cur["price"],
                        "賣出價格":       sell_price,
                        "手續費(買+賣)": buy_fee + sell_fee,
                        "交易稅":        sell_tax,
                        "總成本":        cost,
                        "總獲利":        profit,
                        "報酬率":        f"{rr*100:.2f}%",
                        "備註":          "現股",
                    })
 
                    sell_shares   -= msh
                    cur["shares"] -= msh
                    if cur["shares"] == 0:
                        buy_pool.pop(0)
 
        # 剩餘未賣出的現股庫存
        for b in buy_pool:
            inventory.append({
                "商品名稱":  stock_name,
                "買進時間":  b["date"],
                "持有股數":  b["shares"],
                "買進價格":  b["price"],
                "持股成本":  round(b["price"] * b["shares"]) + round(b["fee_per_share"] * b["shares"]),
                "備註":     "（未實現，需手動填入現價計算損益）",
            })
 
    return pd.DataFrame(realized), pd.DataFrame(inventory)

# ══════════════════════════════════════════════════════════════════════════════
# 📊  彙整報表生成
# ══════════════════════════════════════════════════════════════════════════════
def make_summary_sheets(df_realized: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """依需求產生月報、年報、個股報表"""
    sheets: dict[str, pd.DataFrame] = {}
    if df_realized.empty:
        return sheets

    df = df_realized.copy()
    df["賣出時間_dt"] = pd.to_datetime(df["賣出時間"], format="%Y/%m/%d", errors="coerce")
    df["年份"]  = df["賣出時間_dt"].dt.year
    df["月份"]  = df["賣出時間_dt"].dt.to_period("M").astype(str)

    # 月報
    monthly = (
        df.groupby("月份")
        .agg(
            交易次數=("商品名稱", "count"),
            總損益=("總獲利", "sum"),
            總手續費=("手續費(買+賣)", "sum"),
            總交易稅=("交易稅", "sum"),
        )
        .reset_index()
    )
    monthly["淨損益"] = monthly["總損益"]  # 手續費與稅已在 FIFO 裡扣除
    sheets["月報"] = monthly

    # 年報
    yearly = (
        df.groupby("年份")
        .agg(
            交易次數=("商品名稱", "count"),
            總損益=("總獲利", "sum"),
            總手續費=("手續費(買+賣)", "sum"),
            總交易稅=("交易稅", "sum"),
        )
        .reset_index()
    )
    sheets["年報"] = yearly

    # 個股報表
    stock_summary = (
        df.groupby("商品名稱")
        .agg(
            交易次數=("股數", "count"),
            總交易股數=("股數", "sum"),
            總損益=("總獲利", "sum"),
            平均報酬率=("總獲利", lambda x: f"{(x.sum() / df.loc[x.index, '總成本'].sum() * 100):.2f}%"),
        )
        .reset_index()
        .sort_values("總損益", ascending=False)
    )
    sheets["個股彙整"] = stock_summary

    return sheets


# ══════════════════════════════════════════════════════════════════════════════
# 💾  Excel 輸出（含格式美化）
# ══════════════════════════════════════════════════════════════════════════════
def write_excel(
    output_path: str,
    df_realized:   pd.DataFrame,
    df_unrealized: pd.DataFrame,
    df_all:        pd.DataFrame,
    summary_sheets: dict[str, pd.DataFrame],
) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # 主要工作表
        df_realized.sort_values("賣出時間").to_excel(writer,   sheet_name="實現損益總表", index=False)
        df_unrealized.sort_values("買進時間").to_excel(writer, sheet_name="未實現庫存",   index=False)
        df_all.to_excel(writer,        sheet_name="交易明細歷史", index=False)

        # 彙整報表
        for name, df_s in summary_sheets.items():
            df_s.to_excel(writer, sheet_name=name, index=False)

        # 簡單欄寬自動調整
        for sheet in writer.sheets.values():
            for col in sheet.columns:
                max_len = max(
                    (len(str(cell.value)) for cell in col if cell.value),
                    default=8,
                )
                sheet.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)


# ══════════════════════════════════════════════════════════════════════════════
# 🖨️  終端機摘要列印
# ══════════════════════════════════════════════════════════════════════════════
def print_summary(df_realized: pd.DataFrame, df_unrealized: pd.DataFrame) -> None:
    print("\n" + "═" * 55)
    print("  📊  損益摘要")
    print("═" * 55)

    if df_realized.empty:
        print("  （本次範圍內無已實現損益）")
    else:
        total   = df_realized["總獲利"].sum()
        wins    = (df_realized["總獲利"] > 0).sum()
        losses  = (df_realized["總獲利"] < 0).sum()
        fee     = df_realized["手續費(買+賣)"].sum()
        tax     = df_realized["交易稅"].sum()
        cost    = df_realized["總成本"].sum()
        rr      = total / cost * 100 if cost > 0 else 0

        sign = "+" if total >= 0 else ""
        print(f"  已實現損益：{sign}{total:,.0f} 元  （報酬率 {rr:.2f}%）")
        print(f"  獲利次數：{wins}  虧損次數：{losses}")
        print(f"  已付手續費：{fee:,.0f} 元   已付交易稅：{tax:,.0f} 元")

        # 前三名個股
        top = (
            df_realized.groupby("商品名稱")["總獲利"]
            .sum()
            .sort_values(ascending=False)
            .head(3)
        )
        print("\n  🏆 獲利前三名：")
        for name, val in top.items():
            sign = "+" if val >= 0 else ""
            print(f"     {name}  {sign}{val:,.0f} 元")

    if not df_unrealized.empty:
        total_cost = df_unrealized["持股成本"].sum()
        print(f"\n  📦 未實現庫存：{len(df_unrealized)} 筆，帳面成本 {total_cost:,.0f} 元")
        for _, r in df_unrealized.iterrows():
            print(f"     {r['商品名稱']}  {r['持有股數']} 股 @ {r['買進價格']} 元")

    print("═" * 55 + "\n")


# ══════════════════════════════════════════════════════════════════════════════
# 🚀  主程式
# ══════════════════════════════════════════════════════════════════════════════
def main() -> None:
    args = parse_args()

    # ── 取得密碼 ──────────────────────────────────────────────────────────────
    password = (
        args.password
        or os.getenv("CATHAY_ID")
        or input("請輸入身分證字號（字母大寫）：").strip()
    )

    # ── 決定 PDF 來源清單 ─────────────────────────────────────────────────────
    pdf_paths: list[str] = []

    if args.gmail:
        pdf_paths = fetch_pdfs_from_gmail(
            save_dir   = args.gmail_save_dir,
            start_date = args.gmail_start,
            end_date   = args.gmail_end,
        )
        if not pdf_paths:
            print("❌ 未從 Gmail 取得任何 PDF，程式結束。")
            sys.exit(1)

    elif args.folder:
        folder = Path(args.folder)
        pdf_paths = sorted(str(p) for p in folder.glob("*.pdf"))
        if not pdf_paths:
            print(f"❌ 資料夾 {args.folder} 內無 PDF 檔案。")
            sys.exit(1)
        print(f"📂 找到 {len(pdf_paths)} 個 PDF：{[Path(p).name for p in pdf_paths]}")

    else:
        single = args.pdf or DEFAULT_PDF_PATH
        if not os.path.exists(single):
            print(f"❌ 找不到 PDF：{single}")
            sys.exit(1)
        pdf_paths.append(single)

    # ── 讀取歷史 Excel ────────────────────────────────────────────────────────
    excel_path  = args.output
    history_dates: set[str] = set()
    df_history  = pd.DataFrame()

    if os.path.exists(excel_path):
        try:
            df_history = pd.read_excel(excel_path, sheet_name="交易明細歷史")
            if "時間" in df_history.columns:
                history_dates = set(df_history["時間"].astype(str).unique())
        except Exception as e:
            print(f"⚠️  讀取歷史 Excel 失敗：{e}")

    # ── 逐一解析 PDF ──────────────────────────────────────────────────────────
    new_dfs: list[pd.DataFrame] = []

    for pdf_path in pdf_paths:
        print(f"\n🔓 解析：{Path(pdf_path).name}")
        trade_date, raw_rows = parse_pdf(pdf_path, password)

        if not trade_date:
            print(f"  ⚠️  無法取得日期，略過。")
            continue

        print(f"  📅 對帳單日期：{trade_date}")

        if trade_date in history_dates:
            print(f"  🛑 {trade_date} 已存在歷史紀錄，略過（防止重複）。")
            continue

        if not raw_rows:
            print(f"  ⚠️  無有效交易明細，略過。")
            continue

        df_new = rows_to_df(trade_date, raw_rows)
        print(f"  ✅ 新增 {len(df_new)} 筆交易")
        new_dfs.append(df_new)

    # ── 合併歷史 + 新資料 ─────────────────────────────────────────────────────
    if new_dfs:
        df_all = pd.concat([df_history] + new_dfs, ignore_index=True).drop_duplicates()
    else:
        df_all = df_history
        print("\n ℹ️  沒有新資料需要寫入。")

    if df_all.empty:
        print("❌ 無任何交易資料，結束。")
        sys.exit(0)

    # ── FIFO 計算（全部歷史） ─────────────────────────────────────────────────
    print("\n🧠 執行 FIFO 損益配對…")
    df_realized, df_unrealized = run_fifo(df_all)

    # ── 日期範圍篩選（只影響顯示/輸出，不影響 FIFO 計算邏輯） ──────────────
    df_display = df_realized.copy()

    if args.start:
        df_display = df_display[df_display["賣出時間"] >= args.start]
        print(f"  📅 篩選起始：{args.start}")
    if args.end:
        df_display = df_display[df_display["賣出時間"] <= args.end]
        print(f"  📅 篩選結束：{args.end}")
    if args.stock:
        df_display = df_display[df_display["商品名稱"].str.contains(args.stock)]
        print(f"  🔍 篩選股票：{args.stock}")

    # ── 產生彙整報表 ──────────────────────────────────────────────────────────
    summary_sheets = make_summary_sheets(df_display)

    # ── 終端機摘要 ────────────────────────────────────────────────────────────
    print_summary(df_display, df_unrealized)

    # ── 寫入 Excel ────────────────────────────────────────────────────────────
    if not args.dry_run:
        write_excel(excel_path, df_display, df_unrealized, df_all, summary_sheets)
        print(f"🎉 Excel 已儲存：{excel_path}")
        print(f"   工作表：實現損益總表、未實現庫存、交易明細歷史、月報、年報、個股彙整")
    else:
        print("（Dry-run 模式：未寫入 Excel）")


if __name__ == "__main__":
    main()
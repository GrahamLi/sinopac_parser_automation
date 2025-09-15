import tabula
import pandas as pd
import getpass
import fitz # PyMuPDF
import sys
import re
from datetime import datetime

# --- 設定 ---
PDF_FILE_PATH = "statement.pdf"
OUTPUT_CSV_FILE = "statement_parsed.csv"
# --- 設定結束 ---

def extract_statement_date(doc):
    """從 PDF 文字中嘗試找出帳單年月"""
    year, month = None, None
    for page_num in range(len(doc)):
        page_text = doc[page_num].get_text()
        match = re.search(r'中華民國\s*(\d{3})\s*年\s*(\d{2})\s*月', page_text)
        if match:
            year = int(match.group(1)) + 1911
            month = int(match.group(2))
            return year, month
    # 如果找不到，備用方案
    if doc.metadata and 'creationDate' in doc.metadata:
        date_str = doc.metadata['creationDate']
        match = re.search(r'D:(\d{4})(\d{2})', date_str)
        if match:
            return int(match.group(1)), int(match.group(2))
    return datetime.now().year, datetime.now().month

def main():
    try:
        pdf_password = getpass.getpass(prompt='請輸入您的 PDF 密碼: ')
    except Exception as e:
        print(f"無法讀取密碼: {e}")
        return

    try:
        doc = fitz.open(PDF_FILE_PATH)
        if doc.is_encrypted and not doc.authenticate(pdf_password):
            print("❌ 密碼錯誤或 PDF 無法解密。")
            sys.exit()
        print("✅ PDF 密碼驗證成功！")
        
        statement_year, statement_month = extract_statement_date(doc)
        print(f"偵測到帳單年月為: {statement_year}年 {statement_month}月")
        doc.close()
    except Exception as e:
        print(f"❌ 開啟 PDF 檔案時發生錯誤: {e}")
        return

    print("正在使用 tabula 全頁掃描模式解析 PDF，請稍候...")
    try:
        dfs = tabula.read_pdf(PDF_FILE_PATH, password=pdf_password, pages='all', stream=True)
    except Exception as e:
        print(f"❌ Tabula 解析失敗: {e}")
        return

    if not dfs:
        print("❌ 在 PDF 中找不到任何表格資料。")
        return

    print(f"✅ Tabula 成功解析出 {len(dfs)} 個表格區塊。")
    
    combined_df = pd.concat(dfs, ignore_index=True)
    valid_rows = []
    description_buffer = []

    for index, row in combined_df.iterrows():
        row_text = " ".join([str(v) for v in row.values if pd.notna(v) and 'nan' not in str(v).lower()]).strip()
        if not row_text:
            if description_buffer: description_buffer = []
            continue
        
        # *** 恢復使用您那版更靈活、更強大的正規表示式 ***
        match = re.search(r"(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(?:\d{4}(?:\.0)?\s+)?(.*?)\s*(-?[\d,]+(?:\.\d+)?)$", row_text)

        if match:
            date, post_date, desc_part, amount_str = match.groups()
            
            full_description = " ".join(description_buffer).strip()
            if desc_part:
                full_description = f"{full_description} {desc_part}".strip()
            
            amount = pd.to_numeric(str(amount_str).replace(',', ''), errors='coerce')

            if pd.notna(amount):
                valid_rows.append([date, post_date, full_description, amount])
            
            description_buffer = []
        else:
            if "繳款截止日" not in row_text and "應繳總金額" not in row_text:
                 description_buffer.append(row_text)

    if not valid_rows:
        print("❌ 未能從解析出的資料中篩選出任何有效的交易紀錄。")
        return
        
    print(f"✅ 成功篩選出 {len(valid_rows)} 筆可能的交易紀錄，正在格式化...")
    
    final_df = pd.DataFrame(valid_rows, columns=['原始日期', '附註', '項目', '金額'])
    final_df['金額'] = final_df['金額'].astype(float)
    final_df = final_df[final_df['金額'] > 0].copy()
    final_df['金額'] = final_df['金額'].astype(int)

    def format_date(row):
        try:
            date_str = str(row['原始日期'])
            if re.match(r"^\d{2}/\d{2}$", date_str):
                month, day = map(int, date_str.split('/'))
                
                # *** 全新、更可靠的跨年份判斷邏輯 ***
                year_to_use = statement_year
                # 如果帳單月份是 1月 或 2月，但消費月份是 11月 或 12月，則年份減一
                if statement_month <= 2 and month >= 11:
                    year_to_use -= 1
                return datetime(year_to_use, month, day).date()
            return None
        except:
            return None

    final_df['日期物件'] = final_df.apply(format_date, axis=1)
    final_df.dropna(subset=['日期物件'], inplace=True)

    final_df['年度'] = final_df['日期物件'].apply(lambda x: x.year)
    final_df['月份'] = final_df['日期物件'].apply(lambda x: x.month)
    final_df['日期'] = final_df['日期物件'].apply(lambda x: x.day)
    
    output_df = final_df[['年度', '月份', '日期', '項目', '金額', '附註']]
    output_df = output_df.sort_values(by=['年度', '月份', '日期']).reset_index(drop=True)

    print("\n--- 解析結果預覽 ---")
    print(output_df.to_string())

    try:
        output_df.to_csv(OUTPUT_CSV_FILE, index=False, encoding='utf-8-sig')
        print(f"\n✅ 結果已成功儲存至 {OUTPUT_CSV_FILE}")
    except Exception as e:
        print(f"\n❌ 儲存 CSV 檔案時發生錯誤: {e}")

if __name__ == '__main__':
    main()
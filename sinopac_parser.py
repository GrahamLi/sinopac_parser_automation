import tabula
import pandas as pd
import getpass
import fitz # PyMuPDF
import sys
import re
from datetime import datetime, date, timedelta
import numpy as np
import argparse
import os.path
import base64
import calendar

# --- Google API Imports ---
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- 設定 ---
# 如果修改 SCOPES，請刪除 token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
# --- 設定結束 ---

def _process_gmail_parts_recursive(parts, msg_id, service, year, month):
    for part in parts:
        if part.get('filename'):
            filename = part['filename']
            # Skip the smime.p7s file and other non-relevant files
            if filename == 'smime.p7s' or not filename.lower().endswith('.pdf'):
                continue

            # New, robust filtering logic
            if '帳單' in filename and '繳款聯' not in filename:
                print(f"✅ 找到附件: {filename}")
                if 'data' in part['body']:
                    attachment_data = part['body']['data']
                else:
                    att_id = part['body']['attachmentId']
                    attachment = service.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
                    attachment_data = attachment['data']
                
                file_data = base64.urlsafe_b64decode(attachment_data.encode('UTF-8'))
                
                downloaded_pdf_path = f'sinopac_statement_{year}-{month:02d}.pdf'
                with open(downloaded_pdf_path, 'wb') as f:
                    f.write(file_data)
                print(f"✅ 附件已下載至: {downloaded_pdf_path}")
                return downloaded_pdf_path
        
        # If this part has nested parts, recurse
        if 'parts' in part:
            result = _process_gmail_parts_recursive(part['parts'], msg_id, service, year, month)
            if result:
                return result
    return None

def fetch_latest_bill_from_gmail(year, month):
    """
    連接到 Gmail，根據年份和月份尋找永豐銀行帳單，並下載附件。
    返回下載的 PDF 路徑，如果找不到則返回 None。
    """
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("❌ 找不到 credentials.json 檔案。請確保您已完成 Google API 的設定步驟。" )
                return None
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('gmail', 'v1', credentials=creds)

        start_date = date(year, month, 1)
        _, last_day = calendar.monthrange(year, month)
        end_date = date(year, month, last_day) + timedelta(days=1)
        date_query = f"after:{start_date.strftime('%Y/%m/%d')} before:{end_date.strftime('%Y/%m/%d')}"
        
        query = f'from:(ebillservice@newebill.banksinopac.com.tw) -from:(SpendService@sinopac.com) subject:("電子帳單通知") -subject:("消費通知") has:attachment {date_query}'
        
        results = service.users().messages().list(userId='me', q=query, maxResults=1).execute()
        messages = results.get('messages', [])

        if not messages:
            print(f'❌ 在 Gmail 中找不到 {year} 年 {month} 月的永豐銀行帳單郵件。')
            return None

        msg_id = messages[0]['id']
        message = service.users().messages().get(userId='me', id=msg_id).execute()
        
        downloaded_pdf_path = _process_gmail_parts_recursive(message['payload']['parts'], msg_id, service, year, month)
        if downloaded_pdf_path:
            return downloaded_pdf_path
        
        print(f'❌ 在 {year} 年 {month} 月的郵件中找不到符合條件的附件。')
        return None

    except HttpError as error:
        print(f'An error occurred: {error}')
        return None

def extract_statement_year(doc):
    """從 PDF 文字中嘗試找出帳單年份"""
    for page_num in range(len(doc)):
        page_text = doc[page_num].get_text()
        match = re.search(r'中華民國\s*(\d{3})\s*年', page_text)
        if match:
            roc_year = int(match.group(1))
            return roc_year + 1911
    return datetime.now().year

def main():
    parser = argparse.ArgumentParser(description="解析永豐銀行信用卡帳單，可選從 Gmail 擷取。" )
    parser.add_argument("--year", type=int, help="要從 Gmail 擷取的帳單年份 (例如: 2025)")
    parser.add_argument("--month", type=int, help="要從 Gmail 擷取的帳單月份 (例如: 8)")
    args = parser.parse_args()

    pdf_to_process = None
    current_output_csv_file = None # Declare here

    if args.year and args.month:
        print(f"正在從 Gmail 尋找 {args.year} 年 {args.month} 月的帳單...")
        pdf_to_process = fetch_latest_bill_from_gmail(args.year, args.month)
        if not pdf_to_process:
            return
        current_output_csv_file = f"statement_parsed_{args.year}-{args.month:02d}.csv"
    else:
        print("未指定年份和月份，將嘗試處理本地的 statement.pdf 檔案。" )
        pdf_to_process = "statement.pdf"
        if not os.path.exists(pdf_to_process):
            print(f"❌ 本地模式錯誤: {pdf_to_process} 檔案不存在。" )
            return
        current_output_csv_file = "statement_parsed.csv" # Default for local mode

    try:
        pdf_password = getpass.getpass(prompt=f'請輸入 {os.path.basename(pdf_to_process)} 的 PDF 密碼: ')
    except Exception as e:
        print(f"無法讀取密碼: {e}")
        return

    try:
        doc = fitz.open(pdf_to_process)
        if doc.is_encrypted and not doc.authenticate(pdf_password):
            print("❌ 密碼錯誤或 PDF 無法解密。" )
            sys.exit()
        print("✅ PDF 密碼驗證成功！")
        statement_year = extract_statement_year(doc)
        print(f"偵測到帳單年份為: {statement_year}")
        doc.close()
    except Exception as e:
        print(f"❌ 開啟 PDF 檔案時發生錯誤: {e}")
        return

    print("正在使用 tabula 全頁掃描模式解析 PDF，請稍候...")
    try:
        dfs = tabula.read_pdf(pdf_to_process, password=pdf_password, pages='all', stream=True, guess=False, area=[0, 0, 842, 595])
    except Exception as e:
        print(f"❌ Tabula 解析失敗: {e}")
        print("請確認您的電腦是否已正確安裝 Java。" )
        return

    if not dfs:
        print("❌ 在 PDF 中找不到任何表格資料。" )
        return

    print(f"✅ Tabula 成功解析出 {len(dfs)} 個表格區塊。" )
    
    combined_df = pd.concat(dfs, ignore_index=True)
    valid_rows = []
    description_buffer = []

    for index, row in combined_df.iterrows():
        row_text = " ".join([str(v) for v in row.values if pd.notna(v) and 'nan' not in str(v).lower()]).strip()
        if not row_text:
            if description_buffer:
                description_buffer = []
            continue

        match = re.search(r"(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(?:\d{4}(?:\.0)?\s+)?(.*)$", row_text)

        if match:
            date, post_date, raw_description = match.groups()
            
            # Combine description buffer with raw_description
            full_description = " ".join(description_buffer).strip()
            if raw_description:
                full_description = f"{full_description} {raw_description}".strip()

            # Sanity check (keep this)
            header_keywords = ['帳單說明', '臺幣金額', '總費用', '消費日', '卡號']
            if any(keyword in full_description for keyword in header_keywords):
                description_buffer = []
                continue

            amount = None # Initialize amount

            # Try to find the amount from the end of the full_description
            # This regex looks for a number at the end, optionally preceded by currency codes or dates.
            # It tries to be smart about not picking up dates as amounts.
            # This is a greedy match for the amount at the end.
            amount_match = re.search(r'(-?[\d,]+(?:\.\d+)?)(?:[^\d,.-]*)$', full_description)
            
            if amount_match:
                amount_str = amount_match.group(1)
                amount = pd.to_numeric(amount_str.replace(',', ''), errors='coerce')
                
                # Remove the extracted amount and any trailing non-numeric characters from the description
                full_description = re.sub(r'(-?[\d,]+(?:\.\d+)?)(?:[^\d,.-]*)$', '', full_description).strip()
                
                # --- Start of Gemini Edit ---
                # Check for foreign currency amount in the description and use it if present.
                foreign_match = re.search(r'(USD|US|JPY|EUR)\s+([\d,.-]+)', full_description)
                if foreign_match:
                    foreign_amount_str = foreign_match.group(2)
                    foreign_amount = pd.to_numeric(foreign_amount_str.replace(',', ''), errors='coerce')
                    if pd.notna(foreign_amount):
                        amount = foreign_amount  # Override amount with foreign currency amount
                        # Clean the foreign currency information from the description
                        full_description = re.sub(r'\s*(USD|US|JPY|EUR)\s+[\d,.-]+', '', full_description).strip()
                # --- End of Gemini Edit ---
                
                # Clean up extra spaces after removal
                full_description = re.sub(r'\s+', ' ', full_description).strip()

            if pd.notna(amount):
                valid_rows.append([date, post_date, full_description, amount])
            description_buffer = []
        else:
            # If a row does not match the transaction pattern, it might be a junk row.
            # Clear the buffer if it contains date-like patterns, to prevent carrying over descriptions.
            if re.search(r'\d{2}/\d{2}', row_text):
                description_buffer = []
            else:
                description_buffer.append(row_text)

    if not valid_rows:
        print("❌ 未能從解析出的資料中篩選出任何有效的交易紀錄。" )
        return
        
    print(f"✅ 成功篩選出 {len(valid_rows)} 筆可能的交易紀錄，正在格式化..." )
    
    final_df = pd.DataFrame(valid_rows, columns=['原始日期', '附註', '項目', '金額'])
    
    final_df['金額'] = final_df['金額'].astype(float)
    
    # Identify the main payment (largest negative value) and mark it.
    # This preserves refunds (other negative values) and zero-amount entries.
    if any(final_df['金額'] < 0):
        min_amount = final_df['金額'].min()
        payment_rows = final_df['金額'] == min_amount
        final_df.loc[payment_rows, '項目'] = "永豐自扣已入帳,謝謝!"

    # Create a copy for final output, filtering out the main payment row.
    final_df = final_df[final_df['項目'] != "永豐自扣已入帳,謝謝!"].copy()
    
    # Convert amounts to integers for cleaner output, now that we're done with float comparisons.
    final_df['金額'] = final_df['金額'].astype(int)

    def format_date(row):
        try:
            date_str = str(row['原始日期'])
            if re.match(r"^\d{2}/\d{2}$", date_str):
                month, day = map(int, date_str.split('/'))
                return datetime(statement_year, month, day).date()
            return None
        except:
            return None

    final_df['日期物件'] = final_df.apply(format_date, axis=1)
    final_df.dropna(subset=['日期物件'], inplace=True)

    final_df['年度'] = final_df['日期物件'].apply(lambda x: x.year)
    final_df['月份'] = final_df['日期物件'].apply(lambda x: x.month)
    final_df['日期'] = final_df['日期物件'].apply(lambda x: x.day)
    
    # Step 1: Prepare columns for renaming and new assignments
    final_df['附註1'] = final_df['項目'] # Copy original '項目' content to '附註1'
    final_df.rename(columns={'附註': '附註2'}, inplace=True) # Rename '附註' to '附註2'

    # Step 2: Apply conditional logic to the '項目' column
    # Initialize '項目' with a default value or keep it as is if no condition matches
    # The user implies the original '項目' content should be moved, and the new '項目' is derived.
    # So, we need to create a new '項目' column.
    final_df['新項目'] = '永豐信用卡卡費' # Default value for the new '項目' column

    # Apply conditions based on '附註1'
    final_df.loc[final_df['附註1'].str.contains('優步|和雲|中油', na=False), '新項目'] = '交通'
    final_df.loc[final_df['附註1'].str.contains('凱基', na=False), '新項目'] = '人壽保險費'
    final_df.loc[final_df['附註1'].str.contains('中華電信', na=False), '新項目'] = '電信費'

    # Step 3: Select and order columns for the final output DataFrame
    # Create a copy to avoid SettingWithCopyWarning
    output_df = final_df[['年度', '月份', '日期', '新項目', '金額', '附註1', '附註2']].copy()
    output_df.rename(columns={'新項目': '項目'}, inplace=True) # Rename '新項目' back to '項目' for final output

    output_df = output_df.sort_values(by=['月份', '日期']).reset_index(drop=True)

    # Step 4: Adjust total_row for new column structure
    total_amount = output_df['金額'].sum()
    total_row = pd.DataFrame({
        '年度': [''], '月份': [''], '日期': [''],
        '項目': ['本月總計'], '金額': [total_amount], '附註1': [''], '附註2': [''] # Add new columns
    })
    output_df = pd.concat([output_df, total_row], ignore_index=True)

    print("\n--- 解析結果預覽 ---")
    print(output_df.to_string())

    try:
        output_df.to_csv(current_output_csv_file, index=False, encoding='utf-8-sig')
        print(f"\n✅ 結果已成功儲存至 {current_output_csv_file}")
    except Exception as e:
        print(f"\n❌ 儲存 CSV 檔案時發生錯誤: {e}")

if __name__ == '__main__':
    main()

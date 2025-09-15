永豐銀行信用卡帳單解析器
一個自動化工具，可以從 Gmail 下載永豐銀行信用卡帳單 PDF，並將其解析為結構化的 CSV 格式，方便進行財務管理和記帳。
功能特色
✅ 自動 Gmail 整合 - 直接從 Gmail 下載指定月份的帳單
✅ 智能 PDF 解析 - 使用 Tabula 精確提取表格資料
✅ 多幣別支援 - 自動識別 USD、JPY、EUR 等外幣交易
✅ 交易分類 - 自動將交易分類為交通、保險、電信等類別
✅ 密碼保護 - 支援加密的 PDF 帳單
✅ CSV 輸出 - 生成易於分析的結構化資料
系統需求
必要軟體

Python 3.7+
Java Runtime Environment (JRE) - Tabula 需要

Python 套件
pip install tabula-py pandas PyMuPDF google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client numpy
安裝步驟
1. 克隆或下載程式碼
git clone <repository-url>
cd sinopac-parser
2. 安裝相依套件
pip install -r requirements.txt
3. 設定 Google API 憑證
3.1 建立 Google Cloud Project

前往 Google Cloud Console
建立新專案或選擇現有專案
啟用 Gmail API

3.2 建立 OAuth 2.0 憑證

在 Google Cloud Console 中，前往「APIs & Services」→「Credentials」
點擊「Create Credentials」→「OAuth client ID」
選擇「Desktop application」
下載 JSON 檔案並重新命名為 credentials.json
將 credentials.json 放在程式根目錄

4. 首次執行授權
初次執行時會開啟瀏覽器進行 Google 帳號授權，完成後會產生 token.json 檔案。
使用方法
模式一：從 Gmail 自動下載並解析
# 解析 2025年8月的帳單
python sinopac_parser_local_GeminiCLI.py --year 2025 --month 8

# 解析 2024年12月的帳單
python sinopac_parser_local_GeminiCLI.py --year 2024 --month 12

模式二：解析本地 PDF 檔案
# 將 PDF 檔案命名為 statement.pdf 並放在程式目錄
python sinopac_parser_local_GeminiCLI.py

輸出格式
程式會產生 CSV 檔案，包含以下欄位：
欄位名稱說明範例年度交易年份2025月份交易月份8日期交易日期15項目交易分類交通、永豐信用卡卡費金額交易金額350附註1原始交易描述台北捷運公司附註2過帳日期08/16
自動分類規則
程式會根據交易描述自動分類：

交通 - 包含「優步」、「和雲」、「中油」等關鍵字
人壽保險費 - 包含「凱基」關鍵字
電信費 - 包含「中華電信」關鍵字
永豐信用卡卡費 - 其他未分類交易的預設類別

檔案結構
sinopac-parser/
├── sinopac_parser_local_GeminiCLI.py  # 主程式
├── credentials.json                    # Google API 憑證 (需自行建立)
├── token.json                         # 授權 Token (首次執行後自動產生)
├── statement.pdf                      # 本地模式用的 PDF 檔案
├── statement_parsed.csv              # 本地模式輸出檔案
├── statement_parsed_YYYY-MM.csv      # Gmail模式輸出檔案
└── sinopac_statement_YYYY-MM.pdf     # Gmail模式下載的PDF檔案
常見問題
Q: 執行時出現 "Java not found" 錯誤
A: 請確認已安裝 Java Runtime Environment (JRE)
bash# 檢查 Java 安裝
java -version

# Ubuntu/Debian 安裝
sudo apt install default-jre

# macOS 安裝 (使用 Homebrew)
brew install openjdk
Q: Gmail 授權失敗
A: 請確認：

credentials.json 檔案存在且格式正確
Google Cloud Project 已啟用 Gmail API
OAuth 2.0 憑證類型為「Desktop application」

Q: PDF 密碼錯誤
A: 永豐銀行 PDF 密碼通常是：

身分證字號後4碼 + 生日月日 (MMDD)
例如：身分證 A123456789，生日 1985/03/15，密碼為 678915

Q: 找不到帳單郵件
A: 程式會搜尋符合以下條件的郵件：

寄件者：ebillservice@newebill.banksinopac.com.tw
主旨包含：「電子帳單通知」
有 PDF 附件且檔名包含「帳單」

Q: 解析結果不準確
A: 可能原因：

PDF 格式變更 - 永豐銀行可能更新帳單格式
表格結構異常 - 某些特殊交易可能影響解析
建議檢查原始 PDF 檔案結構

進階設定
自訂分類規則
在程式中修改以下區段來新增分類規則：
python# 在 main() 函數中找到這段程式碼
final_df.loc[final_df['附註1'].str.contains('關鍵字1|關鍵字2'), '新項目'] = '自訂分類'
修改搜尋條件
若需要調整 Gmail 搜尋條件，修改 fetch_latest_bill_from_gmail() 函數中的 query 變數。
授權條款
本專案採用 MIT 授權條款。請自由使用、修改和分發，但請保留原始授權聲明。
免責聲明

本工具僅供個人財務管理使用
請確保遵守永豐銀行的服務條款
作者不對使用本工具導致的任何損失負責
建議定期備份重要資料

貢獻
歡迎提交 Issue 和 Pull Request！
回報問題時請提供：

Python 版本
錯誤訊息截圖
執行環境 (Windows/macOS/Linux)
PDF 帳單格式（請移除敏感資訊）


最後更新：2025年9月
版本：1.0
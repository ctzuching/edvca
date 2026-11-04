# 企業文件新舊版本差異分析助手

Flask 應用：上傳兩份 .docx → Python 段落差異比對 → Gemini 影響分析 → 網頁檢視與 Word 差異報告下載。

## 需求與安裝
- Python 3.10 以上
```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env              # Windows: copy .env.example .env
```
編輯 `.env`：填入 `GEMINI_API_KEY`（於 https://aistudio.google.com/apikey 申請），並用下列指令產生 `FLASK_SECRET_KEY`：
`python -c "import secrets;print(secrets.token_hex(32))"`

## 啟動
```bash
python app.py        # 開啟 http://127.0.0.1:5000
```
正式環境可用 `gunicorn -w 2 app:app`（需設定固定的 `FLASK_SECRET_KEY`；多台主機需共用 `JOBS_DIR`）。

## 測試
```bash
python -m unittest discover -s tests -v
```
測試以 mock 取代 Gemini，不需網路與金鑰；不含真實 Gemini 呼叫。

## 比對邏輯
1. 讀取非空白段落，空白正規化。
2. 以 `difflib.SequenceMatcher` 對段落序列比對，得到相同／新增／刪除／替換區塊。
3. 替換區塊內以動態規劃做「保序最大相似度配對」（字元相似度 ≥ 0.5 才算修改），其餘為刪除或新增。
4. 修改段落再做文字層級比對（英數字以字詞、中文以單字）；過長段落（>4000 字）不標示細部差異。
5. 完全相同內容同時被刪除與新增者加註「疑似段落移動」。

## 安全與隱私
- 僅接受 .docx（副檔名、ZIP 結構、`word/document.xml`、大小、解壓縮大小）；檔名只用於顯示，不用作路徑；文件在記憶體中解析，不寫入磁碟。
- 分析結果與報告存於 `instance/jobs/<uuid>/`，以 session 簽章 cookie 綁定擁有者；下載需驗證 32 位十六進位 ID 與擁有者；逾 `JOB_TTL_MINUTES`（預設 60）自動清除。
- 送交 Gemini 的僅有差異段落（單段上限 1500 字），不含檔名；預設遮蔽電子郵件、身分證字號、手機號碼與長數字串（`MASK_SENSITIVE=false` 可關閉）。**其他個資與機密不會被自動辨識**。
- 日誌只記錄工作代碼、統計與錯誤類型，不記錄文件內容或金鑰。

## 已知限制
見網頁與報告中的「限制與揭露」：僅比對文字段落；表格／圖片／註解／追蹤修訂／頁首頁尾未比對；段落移動無法可靠辨識；不偵測格式變更。

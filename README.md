# 辦公自動化工具 (piccompress)

這是一個用於簡化日常辦公流程的自動化工具，包含：
- PDF 壓縮
- 圖片掃描轉文字 (OCR)
- YouTube 影片轉錄文字
- 照片貼入 Word
- 圖片壓縮介面

## 使用方式
執行 `main.py` 後，依提示選擇功能。
----------------------------
README.md（建議放在專案根目錄）
# 辦公自動化工具（piccompress）

這是一個結合多項日常辦公功能的 Python 自動化工具，整合 PDF、圖片、影片轉文字與壓縮功能，
讓使用者能快速處理常見文書任務。

---

## 🚀 功能列表

| 功能 | 說明 |
|------|------|
| 1️⃣ PDF 壓縮 | 批次壓縮桌面上的 PDF 檔案 |
| 2️⃣ 圖片掃描轉文字 | 使用 Tesseract OCR 將圖片/PDF 轉成文字檔 |
| 3️⃣ YouTube 影片轉文字 | 下載影片並自動轉錄字幕內容 |
| 4️⃣ 照片貼入 Word | 將多張圖片插入 Word 文件 |
| 5️⃣ 圖片壓縮 | 提供圖形介面批次縮圖壓縮 |

---

## 🧩 環境需求

- Python 3.11 或以上版本  
- Git（可選，用於下載此專案）
- 已安裝 Tesseract-OCR（用於影像辨識功能）  
  > Windows 可從 [Tesseract 官方安裝包](https://github.com/UB-Mannheim/tesseract/wiki) 下載。

---

## 💾 安裝教學

### 1️⃣ 下載專案

```bash
git clone https://github.com/HRker/piccompress.git
cd piccompress

2️⃣ 建立虛擬環境
python -m venv .venv

3️⃣ 啟動虛擬環境

Windows

.venv\Scripts\activate


macOS / Linux

source .venv/bin/activate

4️⃣ 安裝套件
pip install -r requirements.txt

▶️ 執行程式
python main.py


啟動後，終端機會出現功能選單：

請選擇功能：
1. PDF壓縮
2. 圖片掃描轉文字
3. YouTube影片轉文字
4. 照片貼入Word
5. 圖片壓縮
0. 退出

📂 專案結構
piccompress/
│
├── main.py               # 主程式進入點
├── src/                  # 各模組功能程式
├── tests/                # 測試程式
├── requirements.txt      # 所需套件列表
├── .gitignore            # 忽略暫存與環境檔案
└── README.md             # 專案說明

🧑‍💻 作者

HRker
📧 Email: (可選填)
📦 GitHub: HRker

📜 授權條款

本專案僅供學術與研究使用，禁止用於商業目的。


---

你可以把這份內容直接貼進 PyCharm 裡新建的 `README.md`，  
然後執行以下三行即可同步上 GitHub 👇

```bash
git add README.md
git commit -m "Update README with installation guide"
git push

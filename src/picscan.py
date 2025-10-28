import os

import sys
import pytesseract
from PIL import Image
from pdf2image import convert_from_path

# 設定 Tesseract 路徑（Windows用戶需要）
# 請根據您的實際安裝路徑修改
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def check_tesseract():
    """檢查 Tesseract 是否正確安裝"""
    try:
        pytesseract.get_tesseract_version()
        print("Tesseract 安裝正確")
        return True
    except Exception as e:
        print(f"Tesseract 錯誤: {str(e)}")
        print("請確認 Tesseract 是否正確安裝，以及路徑設定是否正確")
        return False


def scan_image(file_path):
    """處理圖片文件的文字辨識"""
    try:
        # 開啟並預處理圖片
        image = Image.open(file_path)

        # 嘗試進行文字辨識
        text = pytesseract.image_to_string(image, lang='chi_tra+eng')

        if not text.strip():
            print(f"警告: {file_path} 未能識別出任何文字")
            return "未能識別出文字"

        return text
    except Exception as e:
        print(f"處理圖片時發生錯誤 {file_path}: {str(e)}")
        return f"錯誤: {str(e)}"


def scan_pdf(file_path):
    """處理PDF文件的文字辨識"""
    try:
        # 將PDF轉換為圖片
        pages = convert_from_path(file_path)
        text_result = []

        # 處理每一頁
        for i, page in enumerate(pages):
            print(f"正在處理第 {i + 1} 頁...")
            text = pytesseract.image_to_string(page, lang='chi_tra+eng')
            text_result.append(f"=== 第 {i + 1} 頁 ===\n{text}\n")

        return '\n'.join(text_result)
    except Exception as e:
        print(f"處理PDF時發生錯誤 {file_path}: {str(e)}")
        return f"錯誤: {str(e)}"


def main():
    # 檢查 Tesseract 安裝
    if not check_tesseract():
        return

    # 設定輸入和輸出資料夾
    desktop_path = os.path.expanduser("~/Desktop/picscan")

    # 檢查資料夾是否存在
    if not os.path.exists(desktop_path):
        print(f"建立資料夾: {desktop_path}")
        os.makedirs(desktop_path)

    # 檢查資料夾中是否有文件
    files = [f for f in os.listdir(desktop_path) if os.path.isfile(os.path.join(desktop_path, f))]
    if not files:
        print(f"錯誤: 在 {desktop_path} 中沒有找到任何文件")
        return

    # 處理資料夾中的所有文件
    for filename in files:
        file_path = os.path.join(desktop_path, filename)

        # 跳過輸出文件
        if filename.startswith('output_'):
            continue

        print(f"\n開始處理: {filename}")

        # 根據文件類型進行處理
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
            result = scan_image(file_path)
        elif filename.lower().endswith('.pdf'):
            result = scan_pdf(file_path)
        else:
            print(f"不支援的文件格式: {filename}")
            continue

        # 將結果寫入文字檔
        output_file = os.path.join(desktop_path, f"output_{os.path.splitext(filename)[0]}.txt")
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(result)
            print(f"處理完成: {filename}")
            print(f"結果已保存到: {output_file}")
        except Exception as e:
            print(f"保存結果時發生錯誤: {str(e)}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"程式執行錯誤: {str(e)}")
    finally:
        input("\n按Enter鍵結束程式...")
from src.PDF壓縮 import compress_pdf, format_size, install_ghostscript
from src.picscan import check_tesseract, scan_image, scan_pdf
from src.YT.YT_NEW_transcript import YouTubeTranscriber
from src.photo_to_word.photo_to_word import PhotoToWordApp
from src.picsize import ImageCompressorGUI
from pathlib import Path
import os
import tkinter as tk
from tkinter import messagebox


def main():
    """主程式進入點"""
    print("歡迎使用辦公自動化工具")

    while True:
        print("\n請選擇功能：")
        print("1. PDF壓縮")
        print("2. 圖片掃描轉文字")
        print("3. YouTube影片轉文字")
        print("4. 照片貼入Word")
        print("5. 圖片壓縮")
        print("0. 退出")

        choice = input("請輸入選項（0-5）：")

        if choice == "0":
            break

        elif choice == "1":
            # 檢查 Ghostscript
            if not install_ghostscript():
                continue

            # 取得桌面路徑
            desktop = Path.home() / 'Desktop'
            input_folder = desktop / 'PDF'
            output_folder = desktop / 'PDF_Compressed'

            # 確保輸出資料夾存在
            output_folder.mkdir(parents=True, exist_ok=True)

            if not input_folder.exists():
                print(f"錯誤：找不到輸入資料夾: {input_folder}")
                continue

            # 執行 PDF 壓縮
            pdf_files = list(input_folder.glob('*.pdf'))
            if not pdf_files:
                print("PDF資料夾中沒有PDF檔案")
                continue

            for pdf_file in pdf_files:
                input_path = str(pdf_file)
                output_path = str(output_folder / f"compressed_{pdf_file.name}")
                success, result, original_size, compressed_size = compress_pdf(input_path, output_path)

                if success:
                    print(f"成功壓縮：{pdf_file.name}")
                    print(f"原始大小：{format_size(original_size)}")
                    print(f"壓縮後：{format_size(compressed_size)}")
                    print(f"壓縮率：{result:.1f}%")
                else:
                    print(f"壓縮失敗：{pdf_file.name} - {result}")

        elif choice == "2":
            # 檢查 Tesseract 安裝
            if not check_tesseract():
                print("請先安裝 Tesseract-OCR")
                continue

            # 設定輸入資料夾
            desktop_path = os.path.expanduser("~/Desktop/picscan")

            # 檢查資料夾是否存在
            if not os.path.exists(desktop_path):
                print(f"建立資料夾: {desktop_path}")
                os.makedirs(desktop_path)
                print(f"請將要掃描的圖片或PDF放入：{desktop_path}")
                continue

            # 檢查資料夾中是否有文件
            files = [f for f in os.listdir(desktop_path) if os.path.isfile(os.path.join(desktop_path, f))]
            if not files:
                print(f"錯誤: 在 {desktop_path} 中沒有找到任何文件")
                continue

            # 處理每個文件
            for filename in files:
                if filename.startswith('output_'):
                    continue

                file_path = os.path.join(desktop_path, filename)
                print(f"\n開始處理: {filename}")

                if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                    result = scan_image(file_path)
                elif filename.lower().endswith('.pdf'):
                    result = scan_pdf(file_path)
                else:
                    print(f"不支援的文件格式: {filename}")
                    continue

                # 儲存結果
                output_file = os.path.join(desktop_path, f"output_{os.path.splitext(filename)[0]}.txt")
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(result)
                print(f"處理完成: {filename}")
                print(f"結果已保存到: {output_file}")

        elif choice == "3":
            # 啟動 YouTube 轉文字工具
            app = YouTubeTranscriber()
            app.run()

        elif choice == "4":
            try:
                root = tk.Tk()
                app = PhotoToWordApp(root)
                root.mainloop()
            except Exception as e:
                print(f"照片貼入Word功能發生錯誤: {str(e)}")
                messagebox.showerror("錯誤", f"執行時發生錯誤: {str(e)}")

        elif choice == "5":
            try:
                root = tk.Tk()
                app = ImageCompressorGUI(root)
                root.mainloop()
            except Exception as e:
                print(f"圖片壓縮功能發生錯誤: {str(e)}")
                messagebox.showerror("錯誤", f"執行時發生錯誤: {str(e)}")

        else:
            print("無效的選項，請重新選擇")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"程式執行錯誤: {str(e)}")
        print("錯誤詳細資訊：")
        import traceback

        traceback.print_exc()
    finally:
        input("\n按Enter鍵結束程式...")
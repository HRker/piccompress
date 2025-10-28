import os
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
import subprocess
import sys


def install_ghostscript():
    """檢查是否安裝了 Ghostscript"""
    try:
        # 嘗試執行 gswin64c.exe
        subprocess.run(['gswin64c', '--version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True
    except FileNotFoundError:
        messagebox.showerror("錯誤", "請先安裝 Ghostscript:\n"
                                     "1. 下載網址: https://ghostscript.com/releases/gsdnld.html\n"
                                     "2. 下載 Windows 64bit 版本\n"
                                     "3. 安裝完成後重新執行此程式")
        return False


def compress_pdf(input_path, output_path):
    """
    使用 Ghostscript 壓縮 PDF
    """
    try:
        # Ghostscript 命令
        gs_command = [
            'gswin64c',
            '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.4',
            '-dPDFSETTINGS=/ebook',  # 可選: /screen, /ebook, /printer, /prepress
            '-dNOPAUSE',
            '-dQUIET',
            '-dBATCH',
            f'-sOutputFile={output_path}',
            input_path
        ]

        # 執行命令
        result = subprocess.run(gs_command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if result.returncode == 0:
            original_size = os.path.getsize(input_path)
            compressed_size = os.path.getsize(output_path)
            compression_ratio = (1 - compressed_size / original_size) * 100
            return True, compression_ratio, original_size, compressed_size
        else:
            return False, f"Ghostscript 錯誤: {result.stderr.decode()}", 0, 0

    except Exception as e:
        return False, str(e), 0, 0


def format_size(size):
    """將檔案大小轉換為易讀格式"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024
    return f"{size:.2f} GB"


def main():
    # 檢查 Ghostscript
    if not install_ghostscript():
        return

    # 取得桌面路徑
    desktop = Path.home() / 'Desktop'
    input_folder = desktop / 'PDF'
    output_folder = desktop / 'PDF_Compressed'

    # 確保輸出資料夾存在
    output_folder.mkdir(parents=True, exist_ok=True)

    # 檢查輸入資料夾
    if not input_folder.exists():
        messagebox.showerror("錯誤", f"找不到輸入資料夾: {input_folder}")
        return

    # 取得所有PDF檔案
    pdf_files = list(input_folder.glob('*.pdf'))

    if not pdf_files:
        messagebox.showinfo("提示", "PDF資料夾中沒有PDF檔案")
        return

    # 處理每個檔案
    results = []
    total_original_size = 0
    total_compressed_size = 0

    for pdf_file in pdf_files:
        input_path = str(pdf_file)
        output_path = str(output_folder / f"compressed_{pdf_file.name}")

        success, result, original_size, compressed_size = compress_pdf(input_path, output_path)

        if success:
            total_original_size += original_size
            total_compressed_size += compressed_size
            results.append(
                f"{pdf_file.name}:\n"
                f"原始大小: {format_size(original_size)}\n"
                f"壓縮後: {format_size(compressed_size)}\n"
                f"壓縮率: {result:.1f}%\n"
            )
        else:
            results.append(f"{pdf_file.name}: 壓縮失敗 - {result}\n")

    # 計算總體壓縮率
    if total_original_size > 0:
        total_compression_ratio = (1 - total_compressed_size / total_original_size) * 100
        summary = (
            f"\n總結:\n"
            f"總原始大小: {format_size(total_original_size)}\n"
            f"總壓縮後大小: {format_size(total_compressed_size)}\n"
            f"總體壓縮率: {total_compression_ratio:.1f}%"
        )
        results.append(summary)

    # 顯示結果
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("完成", "".join(results))


if __name__ == "__main__":
    main()
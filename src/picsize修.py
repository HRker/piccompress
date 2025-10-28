import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
import os
import threading


class ImageCompressorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("圖片壓縮工具")
        self.root.geometry("600x600")

        # 建立主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 說明文字框
        help_text = """使用說明：
• quality參數說明（範圍1-95）：
  - 數值越大，圖片品質越好，檔案也越大
  - 數值越小，圖片品質越差，檔案也越小
  - 建議值在60-80之間
  - 可以根據實際需求調整
  - 如果壓縮效果不理想，可以調整參數找到最佳平衡點

• 使用步驟：
  1. 點擊「選擇圖片」選擇要壓縮的圖片（可多選）
  2. 調整壓縮品質數值
  3. 點擊「開始壓縮」
  4. 壓縮後的圖片會存放在原資料夾下的compressed_images資料夾中
"""
        help_label = ttk.Label(main_frame, text=help_text, justify=tk.LEFT,
                               background='#f0f0f0', padding=10)
        help_label.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # 選擇圖片按鈕
        self.selected_files = []
        ttk.Button(main_frame, text="選擇圖片", command=self.browse_files).grid(row=1, column=0, columnspan=3, pady=5)

        # 已選擇檔案數量顯示
        self.files_label = ttk.Label(main_frame, text="尚未選擇圖片")
        self.files_label.grid(row=2, column=0, columnspan=3, pady=5)

        # 品質滑動條
        ttk.Label(main_frame, text="壓縮品質(1-95):").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.quality_var = tk.IntVar(value=70)
        self.quality_slider = ttk.Scale(main_frame, from_=1, to=95, orient=tk.HORIZONTAL,
                                        variable=self.quality_var, length=200)
        self.quality_slider.grid(row=3, column=1, sticky=tk.W, pady=5)
        ttk.Label(main_frame, textvariable=self.quality_var).grid(row=3, column=2, pady=5)

        # 開始壓縮按鈕
        ttk.Button(main_frame, text="開始壓縮", command=self.start_compression).grid(row=4, column=0,
                                                                                     columnspan=3, pady=20)

        # 進度條
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, length=400, mode='determinate',
                                            variable=self.progress_var)
        self.progress_bar.grid(row=5, column=0, columnspan=3, pady=10)

        # 狀態文字
        self.status_text = tk.Text(main_frame, height=10, width=60)
        self.status_text.grid(row=6, column=0, columnspan=3, pady=10)

        # 滾動條
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.grid(row=6, column=3, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)

    def browse_files(self):
        files = filedialog.askopenfilenames(
            title='選擇圖片',
            filetypes=[
                ('圖片檔', '*.jpg;*.jpeg;*.png'),
                ('所有檔案', '*.*')
            ]
        )
        if files:
            self.selected_files = files
            self.files_label.config(text=f"已選擇 {len(files)} 個檔案")

    def compress_image(self, image_path, output_path, quality):
        with Image.open(image_path) as img:
            img.save(output_path, quality=quality, optimize=True)

    def update_status(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def start_compression(self):
        if not self.selected_files:
            messagebox.showerror("錯誤", "請先選擇圖片！")
            return

        # 在新執行緒中執行壓縮
        threading.Thread(target=self.compression_thread, daemon=True).start()

    def compression_thread(self):
        quality = self.quality_var.get()

        # 清空狀態文字
        self.status_text.delete(1.0, tk.END)

        # 建立輸出資料夾
        first_file_dir = os.path.dirname(self.selected_files[0])
        output_folder = os.path.join(first_file_dir, 'compressed_images')
        os.makedirs(output_folder, exist_ok=True)

        total_files = len(self.selected_files)
        self.progress_var.set(0)

        for i, input_path in enumerate(self.selected_files, 1):
            try:
                filename = os.path.basename(input_path)
                output_path = os.path.join(output_folder, filename)

                # 獲取原始檔案大小
                original_size = os.path.getsize(input_path) / 1024  # KB

                # 壓縮圖片
                self.compress_image(input_path, output_path, quality)

                # 獲取壓縮後檔案大小
                compressed_size = os.path.getsize(output_path) / 1024  # KB

                # 更新狀態
                status = (f'處理圖片: {filename}\n'
                          f'原始大小: {original_size:.2f} KB\n'
                          f'壓縮後大小: {compressed_size:.2f} KB\n'
                          f'壓縮率: {((original_size - compressed_size) / original_size * 100):.2f}%\n')
                self.update_status(status)

                # 更新進度條
                self.progress_var.set((i / total_files) * 100)

            except Exception as e:
                self.update_status(f'處理圖片 {filename} 時發生錯誤: {str(e)}')

        self.update_status("\n完成所有圖片壓縮！")
        messagebox.showinfo("完成", "圖片壓縮完成！")


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageCompressorGUI(root)
    root.mainloop()
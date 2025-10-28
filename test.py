import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from datetime import datetime


class ParkingFeeCalculator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("停車費計算程序")
        self.root.geometry("600x400")

        # 創建界面元素
        self.create_widgets()

    def create_widgets(self):
        # 選擇文件按鈕
        self.select_btn = tk.Button(self.root, text="選擇Excel文件", command=self.select_files)
        self.select_btn.pack(pady=10)

        # 顯示選中文件的列表框
        self.file_listbox = tk.Listbox(self.root, width=50, height=10)
        self.file_listbox.pack(pady=10)

        # 輸出路徑選擇
        self.output_btn = tk.Button(self.root, text="選擇輸出位置", command=self.select_output)
        self.output_btn.pack(pady=10)

        # 執行按鈕
        self.process_btn = tk.Button(self.root, text="執行計算", command=self.process_files)
        self.process_btn.pack(pady=10)

        self.selected_files = []
        self.output_path = ""

    def select_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        self.selected_files = files
        self.file_listbox.delete(0, tk.END)
        for file in files:
            self.file_listbox.insert(tk.END, file)

    def select_output(self):
        self.output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

    def calculate_fee(self, entry_time, exit_time):
        # 計算停車時間（分鐘）
        duration = (exit_time - entry_time).total_seconds() / 60

        if duration <= 30:
            return 0
        elif duration <= 60:
            return 30
        else:
            hours = duration / 60
            full_hours = int(hours)
            remaining_minutes = duration % 60

            fee = full_hours * 30
            if remaining_minutes > 30:
                fee += 30
            elif remaining_minutes > 0:
                fee += 15

            return fee

    def process_files(self):
        for file in self.selected_files:
            df = pd.read_excel(file, skiprows=1)  # 跳過第一行

            # 添加新列
            df['驗算金額'] = 0.0
            df['是否相符'] = False

            # 計算每行的費用
            for idx, row in df.iterrows():
                entry_time = pd.to_datetime(row['入場時間'])
                exit_time = pd.to_datetime(row['繳費時間'])
                discount = row['折扣金額']

                calculated_fee = self.calculate_fee(entry_time, exit_time)
                final_fee = calculated_fee - discount

                df.at[idx, '驗算金額'] = final_fee
                df.at[idx, '是否相符'] = abs(final_fee - row['繳費金額']) < 0.01

            # 保存結果
            if self.output_path:
                df.to_excel(self.output_path, index=False)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = ParkingFeeCalculator()
    app.run()
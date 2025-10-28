import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os


class ExcelAnalyzer:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Excel異常資料分析工具")
        self.window.geometry("500x300")

        # 設定視窗樣式
        self.window.configure(bg='#f0f0f0')

        # 建立標題標籤
        title_label = tk.Label(
            self.window,
            text="Excel 異常資料分析工具",
            font=('Arial', 16, 'bold'),
            bg='#f0f0f0',
            pady=20
        )
        title_label.pack()

        # 建立按鈕框架
        button_frame = tk.Frame(self.window, bg='#f0f0f0')
        button_frame.pack(expand=True)

        # 建立按鈕
        self.select_button = tk.Button(
            button_frame,
            text="1. 選擇Excel檔案進行初步分析",
            command=self.analyze_excel,
            width=30,
            height=2,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 10)
        )
        self.select_button.pack(pady=10)

        self.format_button = tk.Button(
            button_frame,
            text="2. 選擇分析結果檔案進行進階整理",
            command=self.format_result,
            width=30,
            height=2,
            bg='#2196F3',
            fg='white',
            font=('Arial', 10)
        )
        self.format_button.pack(pady=10)

        # 建立狀態標籤
        self.status_label = tk.Label(
            self.window,
            text="請選擇要執行的操作",
            bg='#f0f0f0',
            font=('Arial', 9)
        )
        self.status_label.pack(pady=20)

        self.window.mainloop()

    def analyze_excel(self):
        try:
            self.status_label.config(text="正在選擇檔案...")
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )

            if file_path:
                self.status_label.config(text="正在分析資料...")
                # 讀取Excel檔案
                df = pd.read_excel(file_path)

                # 找出統編重複的資料
                duplicates = df[df.duplicated(subset=['營利事業統一編號'], keep=False)]
                result = pd.DataFrame()

                # 檢查每個重複的統編
                for tax_id in duplicates['營利事業統一編號'].unique():
                    temp_df = df[df['營利事業統一編號'] == tax_id]
                    # 如果同一統編有不同的受款人名稱
                    if len(temp_df['受款人名稱'].unique()) > 1:
                        result = pd.concat([result, temp_df])

                if not result.empty:
                    # 建立輸出檔案路徑
                    output_path = os.path.splitext(file_path)[0] + '_分析結果.xlsx'
                    # 儲存結果
                    with pd.ExcelWriter(output_path) as writer:
                        result.to_excel(writer, sheet_name='異常資料', index=False)

                    self.status_label.config(text=f"分析完成！結果已儲存至：{os.path.basename(output_path)}")
                    messagebox.showinfo("完成", f"初步分析完成！\n結果已儲存至：{output_path}")
                else:
                    self.status_label.config(text="分析完成，沒有發現異常資料！")
                    messagebox.showinfo("完成", "沒有發現異常資料！")

        except Exception as e:
            self.status_label.config(text="處理過程發生錯誤！")
            messagebox.showerror("錯誤", f"處理過程發生錯誤：{str(e)}")

    def format_result(self):
        try:
            self.status_label.config(text="正在選擇檔案...")
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )

            if file_path:
                self.status_label.config(text="正在進行進階整理...")
                # 讀取分析結果檔案
                df = pd.read_excel(file_path)

                # 先篩選單筆金額小於15萬的資料
                df = df[df['支付金額'] < 150000]

                # 計算每個統編的總金額
                total_amounts = df.groupby('營利事業統一編號')['支付金額'].sum()

                # 找出總金額小於15萬的統編
                valid_ids = total_amounts[total_amounts < 150000].index

                # 只保留總金額小於15萬的統編資料
                df = df[df['營利事業統一編號'].isin(valid_ids)]

                # 建立新的DataFrame，保持原始格式
                formatted_df = pd.DataFrame({
                    '機關名稱': df['機關名稱'],
                    '營利事業統一編號': df['營利事業統一編號'],
                    '受款人名稱': df['受款人名稱'],
                    '筆數': df.groupby('受款人名稱').cumcount() + 1,
                    '加總金額': df['支付金額']
                })

                # 排序結果
                formatted_df = formatted_df.sort_values(
                    ['機關名稱', '營利事業統一編號', '受款人名稱']
                )

                # 建立輸出檔案路徑
                output_path = os.path.splitext(file_path)[0] + '_進階整理.xlsx'

                # 儲存結果
                with pd.ExcelWriter(output_path) as writer:
                    formatted_df.to_excel(writer, sheet_name='整理結果', index=False)

                self.status_label.config(text=f"進階整理完成！結果已儲存至：{os.path.basename(output_path)}")
                messagebox.showinfo("完成", f"進階整理完成！\n結果已儲存至：{output_path}")

        except Exception as e:
            self.status_label.config(text="處理過程發生錯誤！")
            messagebox.showerror("錯誤", f"處理過程發生錯誤：{str(e)}")


if __name__ == "__main__":
    app = ExcelAnalyzer()
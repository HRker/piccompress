import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


class ExcelAnalyzer:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Excel分析工具")
        self.window.geometry("400x250")

        # 建立按鈕
        self.select_button = tk.Button(
            self.window,
            text="選擇Excel檔案進行初步分析",
            command=self.analyze_excel
        )
        self.select_button.pack(pady=20)

        self.format_button = tk.Button(
            self.window,
            text="選擇分析結果檔案進行進階整理",
            command=self.format_result
        )
        self.format_button.pack(pady=20)

        self.window.mainloop()

    def analyze_excel(self):
        # 原本的分析邏輯保持不變
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )

            if file_path:
                df = pd.read_excel(file_path)
                duplicates = df[df.duplicated(subset=['營利事業統一編號'], keep=False)]
                result = pd.DataFrame()

                for tax_id in duplicates['營利事業統一編號'].unique():
                    temp_df = df[df['營利事業統一編號'] == tax_id]
                    if len(temp_df['受款人名稱'].unique()) > 1:
                        result = pd.concat([result, temp_df])

                if not result.empty:
                    output_path = file_path.rsplit('.', 1)[0] + '_分析結果.xlsx'
                    with pd.ExcelWriter(output_path) as writer:
                        result.to_excel(writer, sheet_name='異常資料', index=False)
                    messagebox.showinfo("完成", f"初步分析完成！\n結果已儲存至：{output_path}")
                else:
                    messagebox.showinfo("完成", "沒有發現異常資料！")

        except Exception as e:
            messagebox.showerror("錯誤", f"處理過程發生錯誤：{str(e)}")

    def format_result(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )

            if file_path:
                # 讀取分析結果檔案
                df = pd.read_excel(file_path)

                # 篩選金額小於15萬的資料
                df = df[df['支付金額'] < 150000]

                # 建立新的DataFrame，只保留需要的欄位
                formatted_df = pd.DataFrame({
                    '機關名稱': df['機關名稱'],
                    '營利事業統一編號': df['營利事業統一編號'],
                    '受款人名稱': df['受款人名稱'],
                    '筆數': 1,  # 每筆資料計為1筆
                    '加總金額': df['支付金額']
                })

                # 根據機關名稱和營利事業統一編號分組，並計算筆數和金額總和
                grouped_df = formatted_df.groupby(
                    ['機關名稱', '營利事業統一編號', '受款人名稱']
                ).agg({
                    '筆數': 'sum',
                    '加總金額': 'sum'
                }).reset_index()

                # 儲存結果
                output_path = file_path.rsplit('.', 1)[0] + '_進階整理.xlsx'
                with pd.ExcelWriter(output_path) as writer:
                    grouped_df.to_excel(writer, sheet_name='整理結果', index=False)

                messagebox.showinfo("完成", f"進階整理完成！\n結果已儲存至：{output_path}")

        except Exception as e:
            messagebox.showerror("錯誤", f"處理過程發生錯誤：{str(e)}")


if __name__ == "__main__":
    app = ExcelAnalyzer()
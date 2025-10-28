import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import messagebox


class ExcelAnalyzer:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Excel分析工具")
        self.window.geometry("400x200")

        # 建立按鈕
        self.select_button = tk.Button(
            self.window,
            text="選擇Excel檔案",
            command=self.analyze_excel
        )
        self.select_button.pack(pady=20)

        self.window.mainloop()

    def analyze_excel(self):
        try:
            # 選擇檔案
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )

            if file_path:
                # 讀取Excel檔案
                df = pd.read_excel(file_path)

                # 找出統編相同但名稱不同的資料
                duplicates = df[df.duplicated(subset=['營利事業統一編號'], keep=False)]
                result = pd.DataFrame()

                # 檢查每個重複的統編
                for tax_id in duplicates['營利事業統一編號'].unique():
                    temp_df = df[df['營利事業統一編號'] == tax_id]
                    # 如果同一統編有不同的受款人名稱
                    if len(temp_df['受款人名稱'].unique()) > 1:
                        result = pd.concat([result, temp_df])

                if not result.empty:
                    # 建立新的Excel檔案
                    output_path = file_path.rsplit('.', 1)[0] + '_分析結果.xlsx'
                    with pd.ExcelWriter(output_path) as writer:
                        result.to_excel(writer, sheet_name='異常資料', index=False)

                    messagebox.showinfo("完成", f"分析完成！\n結果已儲存至：{output_path}")
                else:
                    messagebox.showinfo("完成", "沒有發現異常資料！")

        except Exception as e:
            messagebox.showerror("錯誤", f"處理過程發生錯誤：{str(e)}")


if __name__ == "__main__":
    app = ExcelAnalyzer()
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os


class ExcelAnalyzer:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Excel異常資料分析工具")
        self.window.geometry("500x450")

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

        # 建立支出用途關鍵字標籤和輸入框
        purpose_label = tk.Label(
            self.window,
            text="排除支出用途包含:",
            bg='#f0f0f0',
            font=('Arial', 10)
        )
        purpose_label.pack()

        self.purpose_entry = tk.Entry(
            self.window,
            width=40,
            font=('Arial', 10)
        )
        self.purpose_entry.pack(pady=(5, 15))

        # 建立受款人關鍵字標籤和輸入框
        keyword_label = tk.Label(
            self.window,
            text="排除受款人名稱包含:",
            bg='#f0f0f0',
            font=('Arial', 10)
        )
        keyword_label.pack()

        self.keyword_entry = tk.Entry(
            self.window,
            width=40,
            font=('Arial', 10)
        )
        self.keyword_entry.pack(pady=(5, 20))

        # 建立按鈕
        self.select_button = tk.Button(
            self.window,
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
            self.window,
            text="2. 選擇分析結果檔案進行進階整理",
            command=self.format_result,
            width=30,
            height=2,
            bg='#2196F3',
            fg='white',
            font=('Arial', 10)
        )
        self.format_button.pack(pady=10)

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

                # 依機關名稱分組
                result = pd.DataFrame()
                for org_name in df['機關名稱'].unique():
                    org_df = df[df['機關名稱'] == org_name]

                    # 找出統編重複的資料
                    duplicates = org_df[org_df.duplicated(subset=['營利事業統一編號'], keep=False)]

                    # 檢查每個重複的統編
                    for tax_id in duplicates['營利事業統一編號'].unique():
                        temp_df = org_df[org_df['營利事業統一編號'] == tax_id]
                        # 如果同一統編有不同的受款人名稱
                        if len(temp_df['受款人名稱'].unique()) > 1:
                            result = pd.concat([result, temp_df])

                if not result.empty:
                    try:
                        output_path = os.path.splitext(file_path)[0] + '_分析結果.xlsx'
                        # 檢查檔案是否存在並可寫入
                        if os.path.exists(output_path):
                            try:
                                with open(output_path, 'a'):
                                    pass
                            except PermissionError:
                                base, ext = os.path.splitext(output_path)
                                i = 1
                                while True:
                                    new_path = f"{base}_{i}{ext}"
                                    if not os.path.exists(new_path):
                                        output_path = new_path
                                        break
                                    i += 1

                        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                            result.to_excel(writer, sheet_name='異常資料', index=False)

                        self.status_label.config(text=f"分析完成！結果已儲存至：{os.path.basename(output_path)}")
                        messagebox.showinfo("完成", f"初步分析完成！\n結果已儲存至：{output_path}")
                    except PermissionError:
                        self.status_label.config(text="無法儲存檔案，請確認檔案未被其他程式開啟！")
                        messagebox.showerror("錯誤", "無法儲存檔案，請確認：\n1. 檔案未被其他程式開啟\n2. 您有權限寫入該位置")
                else:
                    self.status_label.config(text="分析完成，沒有發現異常資料！")
                    messagebox.showinfo("完成", "沒有發現異常資料！")

        except Exception as e:
            self.status_label.config(text="處理過程發生錯誤！")
            messagebox.showerror("錯誤", f"處理過程發生錯誤：{str(e)}")

    def format_result(self):
        try:
            # 獲取關鍵字並處理（以逗號分隔）
            purpose_keywords = [k.strip() for k in self.purpose_entry.get().split(',') if k.strip()]
            payee_keywords = [k.strip() for k in self.keyword_entry.get().split(',') if k.strip()]

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

                # 如果有輸入支出用途關鍵字，排除相關資料
                if purpose_keywords:
                    mask = df['支出用途'].apply(lambda x: any(keyword in str(x) for keyword in purpose_keywords))
                    df = df[~mask]

                # 如果有輸入受款人關鍵字，排除相關統編的所有資料
                if payee_keywords:
                    mask = df['受款人名稱'].apply(lambda x: any(keyword in str(x) for keyword in payee_keywords))
                    exclude_tax_ids = df[mask]['營利事業統一編號'].unique()
                    df = df[~df['營利事業統一編號'].isin(exclude_tax_ids)]

                # 計算每個統編下不同受款人的總金額
                result_list = []
                for org_name in df['機關名稱'].unique():
                    org_df = df[df['機關名稱'] == org_name]

                    for tax_id in org_df['營利事業統一編號'].unique():
                        tax_df = org_df[org_df['營利事業統一編號'] == tax_id]
                        total_amount = tax_df['支付金額'].sum()

                        # 只處理總金額大於15萬的統編
                        if total_amount > 150000:
                            for payee in tax_df['受款人名稱'].unique():
                                payee_df = tax_df[tax_df['受款人名稱'] == payee]
                                result_list.append({
                                    '機關名稱': org_name,
                                    '營利事業統一編號': tax_id,
                                    '受款人名稱': payee,
                                    '筆數': len(payee_df),
                                    '加總金額': payee_df['支付金額'].sum()
                                })

                if result_list:
                    try:
                        # 建立結果DataFrame
                        result_df = pd.DataFrame(result_list)
                        result_df = result_df.sort_values(['機關名稱', '營利事業統一編號', '受款人名稱'])

                        # 儲存結果
                        output_path = os.path.splitext(file_path)[0] + '_進階整理.xlsx'
                        # 檢查檔案是否存在並可寫入
                        if os.path.exists(output_path):
                            try:
                                with open(output_path, 'a'):
                                    pass
                            except PermissionError:
                                base, ext = os.path.splitext(output_path)
                                i = 1
                                while True:
                                    new_path = f"{base}_{i}{ext}"
                                    if not os.path.exists(new_path):
                                        output_path = new_path
                                        break
                                    i += 1

                        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                            result_df.to_excel(writer, sheet_name='整理結果', index=False)

                        self.status_label.config(text=f"進階整理完成！結果已儲存至：{os.path.basename(output_path)}")
                        messagebox.showinfo("完成", f"進階整理完成！\n結果已儲存至：{output_path}")
                    except PermissionError:
                        self.status_label.config(text="無法儲存檔案，請確認檔案未被其他程式開啟！")
                        messagebox.showerror("錯誤", "無法儲存檔案，請確認：\n1. 檔案未被其他程式開啟\n2. 您有權限寫入該位置")
                else:
                    self.status_label.config(text="沒有符合條件的資料！")
                    messagebox.showinfo("完成", "沒有符合條件的資料！")

        except Exception as e:
            self.status_label.config(text="處理過程發生錯誤！")
            messagebox.showerror("錯誤", f"處理過程發生錯誤：{str(e)}")


if __name__ == "__main__":
    app = ExcelAnalyzer()
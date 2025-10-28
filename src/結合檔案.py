import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import gc  # 用於垃圾回收
import sys  # 用於錯誤輸出
import re  # 用於正則表達式處理


class FileMerger:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("檔案合併工具")
        self.window.geometry("800x700")  # 減少高度
        self.window.minsize(800, 700)  # 設定最小尺寸

        # 檔案路徑變數
        self.file1_path = ""
        self.file2_path = ""

        # 建立日誌檔案
        self.setup_logging()

        # 建立GUI元件
        self.create_widgets()

        # 設定視窗樣式
        self.style_window()

    def setup_logging(self):
        # 建立日誌檔案
        self.log_file = f'merge_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt'
        # 重導向標準輸出和錯誤輸出到檔案
        sys.stdout = open(self.log_file, 'w', encoding='utf-8')
        sys.stderr = sys.stdout

    def log_message(self, message):
        # 在文字區域中添加訊息
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)  # 自動捲動到最新的訊息
        self.window.update()

        # 同時也寫入到日誌檔案
        print(log_entry, end='')
        sys.stdout.flush()

    def style_window(self):
        self.window.configure(bg='#f0f0f0')
        style = ttk.Style()
        style.configure('TButton', padding=5)
        style.configure('TLabel', background='#f0f0f0')

    def select_file1(self):
        """選擇第一個檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇第一個檔案",
            filetypes=[
                ("Excel檔案", "*.xlsx"),
                ("CSV檔案", "*.csv"),
                ("所有檔案", "*.*")
            ]
        )
        if file_path:
            self.file1_path = file_path
            self.file1_label.config(text=os.path.basename(file_path))
            self.log_message(f"已選擇第一個檔案：{file_path}")

    def select_file2(self):
        """選擇第二個檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇第二個檔案",
            filetypes=[
                ("Excel檔案", "*.xlsx"),
                ("CSV檔案", "*.csv"),
                ("所有檔案", "*.*")
            ]
        )
        if file_path:
            self.file2_path = file_path
            self.file2_label.config(text=os.path.basename(file_path))
            self.log_message(f"已選擇第二個檔案：{file_path}")

    def clear_selection(self):
        """清除已選擇的檔案"""
        self.file1_path = ""
        self.file2_path = ""
        self.file1_label.config(text="未選擇檔案")
        self.file2_label.config(text="未選擇檔案")
        self.update_progress(0, "", "")
        self.log_message("已清除檔案選擇")

    def update_progress(self, value, message="", detail=""):
        """更新進度條和訊息"""
        self.progress_bar["value"] = value
        self.progress_label.config(text=message)
        self.detail_label.config(text=detail)
        self.window.update()

    def create_widgets(self):
        # 建立主框架
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 標題
        title_label = ttk.Label(
            main_frame,
            text="檔案合併工具",
            font=('Helvetica', 16, 'bold')
        )
        title_label.pack(pady=10)

        # 說明文字
        instruction_text = """
        使用說明：
        1. 選擇第一個主要檔案（必須包含「機關名稱」、「傳票編號」和「支付金額」欄位）
        2. 選擇第二個合併檔案（必須包含「機關名稱」、「傳票編號」、「支出用途」和「金額」欄位）
        3. 點擊「合併檔案」開始處理

        注意：
        - 處理大檔案時可能需要較長時間，請耐心等待
        - 系統會自動處理傳票編號的格式（包含科學記號格式和文字）
        - 相同傳票編號的支出用途會自動整合
        """
        instruction_label = ttk.Label(
            main_frame,
            text=instruction_text,
            justify=tk.LEFT,
            wraplength=700
        )
        instruction_label.pack(pady=10, padx=5, anchor='w')

        # 第一個檔案選擇區域
        file1_frame = ttk.LabelFrame(main_frame, text="第一個檔案（主檔）", padding="5")
        file1_frame.pack(fill=tk.X, padx=5, pady=5)

        self.file1_button = ttk.Button(
            file1_frame,
            text="選擇檔案",
            command=self.select_file1,
            width=15
        )
        self.file1_button.pack(side=tk.LEFT, padx=5)

        self.file1_label = ttk.Label(file1_frame, text="未選擇檔案")
        self.file1_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 第二個檔案選擇區域
        file2_frame = ttk.LabelFrame(main_frame, text="第二個檔案（合併檔）", padding="5")
        file2_frame.pack(fill=tk.X, padx=5, pady=5)

        self.file2_button = ttk.Button(
            file2_frame,
            text="選擇檔案",
            command=self.select_file2,
            width=15
        )
        self.file2_button.pack(side=tk.LEFT, padx=5)

        self.file2_label = ttk.Label(file2_frame, text="未選擇檔案")
        self.file2_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 進度條框架
        progress_frame = ttk.LabelFrame(main_frame, text="處理進度", padding="5")
        progress_frame.pack(fill=tk.X, padx=5, pady=5)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            length=300,
            mode="determinate"
        )
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)

        self.progress_label = ttk.Label(progress_frame, text="")
        self.progress_label.pack(pady=5)

        self.detail_label = ttk.Label(progress_frame, text="")
        self.detail_label.pack(pady=5)

        # 按鈕框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10, fill=tk.X)

        # 將按鈕置中
        button_container = ttk.Frame(button_frame)
        button_container.pack(anchor=tk.CENTER)

        self.merge_button = ttk.Button(
            button_container,
            text="合併檔案",
            command=self.merge_files,
            width=15
        )
        self.merge_button.pack(side=tk.LEFT, padx=5)

        self.clear_button = ttk.Button(
            button_container,
            text="清除選擇",
            command=self.clear_selection,
            width=15
        )
        self.clear_button.pack(side=tk.LEFT, padx=5)

        # 日誌顯示區域
        log_frame = ttk.LabelFrame(main_frame, text="執行日誌", padding="5")
        log_frame.pack(fill=tk.BOTH, padx=5, pady=5)

        # 建立文字區域和捲軸，設定固定高度
        self.log_text = tk.Text(log_frame, height=6, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # 放置文字區域和捲軸
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, padx=(5, 0), pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

    def convert_voucher_no(self, x):
        try:
            if pd.isna(x):
                return x
            # 將值轉換為字串
            x = str(x)
            # 如果包含科學記號，先處理
            if 'E' in x.upper():
                x = str(int(float(x)))
            # 移除所有非數字字元（包含"新普通公"等文字）
            numbers = re.findall(r'\d+', x)
            return ''.join(numbers) if numbers else x
        except:
            return str(x)

    def merge_files(self):
        if not self.file1_path or not self.file2_path:
            messagebox.showerror("錯誤", "請選擇兩個檔案")
            return

        try:
            # 檢查檔案是否存在
            if not os.path.exists(self.file1_path):
                messagebox.showerror("錯誤", f"找不到第一個檔案：{self.file1_path}")
                return
            if not os.path.exists(self.file2_path):
                messagebox.showerror("錯誤", f"找不到第二個檔案：{self.file2_path}")
                return

            self.update_progress(20, "讀取第一個檔案中...", "正在載入資料")

            # 讀取第一個檔案
            try:
                if self.file1_path.lower().endswith('.xlsx'):
                    df1 = pd.read_excel(
                        self.file1_path,
                        dtype={'機關名稱': str},
                        converters={'傳票編號': self.convert_voucher_no}
                    )
                else:
                    df1 = pd.read_csv(
                        self.file1_path,
                        encoding='utf-8-sig',
                        dtype={'機關名稱': str},
                        converters={'傳票編號': self.convert_voucher_no}
                    )
            except Exception as e:
                self.log_message(f"讀取第一個檔案時發生錯誤：{str(e)}")
                try:
                    df1 = pd.read_csv(
                        self.file1_path,
                        encoding='cp950',
                        dtype={'機關名稱': str},
                        converters={'傳票編號': self.convert_voucher_no}
                    )
                except Exception as e:
                    raise Exception(f"無法讀取第一個檔案：{str(e)}")

            self.log_message(f"第一個檔案欄位：{df1.columns.tolist()}")
            self.log_message(f"第一個檔案資料筆數：{len(df1)}")

            self.update_progress(40, "讀取第二個檔案中...", "正在載入資料")

            # 讀取第二個檔案
            try:
                if self.file2_path.lower().endswith('.xlsx'):
                    df2 = pd.read_excel(
                        self.file2_path,
                        dtype={'機關名稱': str},
                        converters={'傳票編號': self.convert_voucher_no}
                    )
                else:
                    df2 = pd.read_csv(
                        self.file2_path,
                        encoding='utf-8-sig',
                        dtype={'機關名稱': str},
                        converters={'傳票編號': self.convert_voucher_no}
                    )
            except Exception as e:
                self.log_message(f"讀取第二個檔案時發生錯誤：{str(e)}")
                try:
                    df2 = pd.read_csv(
                        self.file2_path,
                        encoding='cp950',
                        dtype={'機關名稱': str},
                        converters={'傳票編號': self.convert_voucher_no}
                    )
                except Exception as e:
                    raise Exception(f"無法讀取第二個檔案：{str(e)}")

            self.log_message(f"第二個檔案欄位：{df2.columns.tolist()}")
            self.log_message(f"第二個檔案資料筆數：{len(df2)}")

            # 建立用於比對的欄位
            self.log_message("處理傳票編號格式...")
            df1['傳票編號_比對用'] = df1['傳票編號'].apply(self.convert_voucher_no)
            df2['傳票編號_比對用'] = df2['傳票編號'].apply(self.convert_voucher_no)

            # 顯示處理後的範例
            self.log_message("傳票編號處理範例：")
            self.log_message(f"第一個檔案：\n{df1[['傳票編號', '傳票編號_比對用']].head()}")
            self.log_message(f"第二個檔案：\n{df2[['傳票編號', '傳票編號_比對用']].head()}")

            # 合併檔案
            self.update_progress(70, "合併檔案中...", "正在處理資料")

            # 進行合併
            merged_df = pd.merge(
                df1,
                df2[['傳票編號_比對用', '支出用途']],
                on='傳票編號_比對用',
                how='left'
            )

            # 移除比對用欄位
            merged_df = merged_df.drop('傳票編號_比對用', axis=1)

            # 檢查未成功合併的記錄
            unmatched_records = merged_df[merged_df['支出用途'].isna()].copy()
            if len(unmatched_records) > 0:
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                unmatched_file = f'未匹配記錄_{current_time}.xlsx'
                unmatched_records.to_excel(unmatched_file, index=False)

                if not messagebox.askyesno("警告",
                                           f"發現 {len(unmatched_records)} 筆記錄未能找到對應的支出用途\n"
                                           f"未匹配記錄已儲存至：{unmatched_file}\n\n"
                                           "是否仍要繼續？"):
                    self.update_progress(0, "", "")
                    return

            # 將未匹配的支出用途填入預設值
            merged_df['支出用途'] = merged_df['支出用途'].fillna('未找到對應支出用途')

            # 釋放記憶體
            del df1, df2
            gc.collect()

            self.update_progress(80, "準備儲存...", "正在準備儲存檔案")

            # 取得當前時間作為檔名一部分
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")

            # 建立儲存檔案的對話框
            save_path = filedialog.asksaveasfilename(
                initialfile=f'合併檔案_{current_time}',
                defaultextension='.xlsx',
                filetypes=[
                    ("Excel檔案", "*.xlsx"),
                    ("CSV檔案", "*.csv"),
                    ("所有檔案", "*.*")
                ]
            )

            if save_path:
                self.update_progress(90, "儲存檔案中...", "正在寫入檔案")
                total_rows = len(merged_df)

                if save_path.endswith('.xlsx'):
                    writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
                    merged_df.to_excel(writer, sheet_name='合併結果', index=False)

                    # 設定格式
                    worksheet = writer.sheets['合併結果']

                    # 設定欄寬
                    for idx, col in enumerate(merged_df.columns):
                        max_length = max(
                            merged_df[col].astype(str).str.len().max(),
                            len(str(col))
                        )
                        worksheet.set_column(idx, idx, min(max_length + 2, 50))

                    # 凍結首行
                    worksheet.freeze_panes(1, 0)
                    writer.close()
                else:
                    merged_df.to_csv(save_path, index=False, encoding='utf-8-sig')

                self.update_progress(100, "完成！", f"總計處理 {total_rows:,} 筆資料")

                messagebox.showinfo(
                    "成功",
                    f"檔案合併完成！\n儲存位置：{save_path}\n總筆數：{total_rows:,}筆"
                )

                if messagebox.askyesno("確認", "是否要開啟檔案所在資料夾？"):
                    os.startfile(os.path.dirname(save_path))
            else:
                self.update_progress(0, "", "")

        except Exception as e:
            self.log_message(f"發生錯誤：{str(e)}")
            self.update_progress(0, "發生錯誤", str(e))
            messagebox.showerror("錯誤", f"處理檔案時發生錯誤：{str(e)}")

    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    app = FileMerger()
    app.run()
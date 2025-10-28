import yt_dlp
import whisper
import os
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import threading
import warnings
import signal
import sys
import time
import gc
import pydub

# 忽略警告
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)


class VideoTranscriber:
    def __init__(self):
        # 設定 FFmpeg 路徑
        self.FFMPEG_PATH = r"C:\Program Files\ffmpeg\bin"

        # 檢查 FFmpeg 是否存在
        ffmpeg_exe = os.path.join(self.FFMPEG_PATH, "ffmpeg.exe")
        if not os.path.exists(ffmpeg_exe):
            messagebox.showerror("錯誤", f"找不到 FFmpeg，請確認是否已安裝在：{ffmpeg_exe}")
            sys.exit(1)

        os.environ["PATH"] += os.pathsep + self.FFMPEG_PATH

        # 建立主視窗
        self.window = tk.Tk()
        self.window.title("音訊轉文字工具")
        self.window.geometry("600x500")

        # 初始化變數
        self.selected_mode = tk.StringVar(value="youtube")
        self.mp4_path = tk.StringVar()
        self.url_var = tk.StringVar()
        self.current_thread = None
        self.is_processing = False
        self.start_time = None

        # 建立介面
        self.create_widgets()

        # 設定關閉視窗處理
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        # 模式選擇框架
        mode_frame = ttk.LabelFrame(self.window, text="選擇模式", padding=10)
        mode_frame.pack(fill="x", padx=10, pady=5)

        # 單選按鈕
        ttk.Radiobutton(mode_frame, text="YouTube 影片", value="youtube",
                        variable=self.selected_mode, command=self.toggle_mode).pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="本地音訊檔案", value="mp4",
                        variable=self.selected_mode, command=self.toggle_mode).pack(side="left", padx=5)

        # YouTube 框架
        self.youtube_frame = ttk.LabelFrame(self.window, text="YouTube 網址", padding=10)
        self.youtube_frame.pack(fill="x", padx=10, pady=5)

        self.url_entry = ttk.Entry(self.youtube_frame, textvariable=self.url_var, width=50)
        self.url_entry.pack(side="left", padx=5)

        # 音訊框架
        self.mp4_frame = ttk.LabelFrame(self.window, text="選擇音訊檔案 (支援 m4a, mp4, mp3, wav)", padding=10)
        self.mp4_frame.pack(fill="x", padx=10, pady=5)

        self.mp4_entry = ttk.Entry(self.mp4_frame, textvariable=self.mp4_path, width=50)
        self.mp4_entry.pack(side="left", padx=5)

        self.browse_btn = ttk.Button(self.mp4_frame, text="瀏覽", command=self.browse_mp4)
        self.browse_btn.pack(side="left", padx=5)

        # 控制按鈕框架
        control_frame = ttk.Frame(self.window)
        control_frame.pack(fill="x", padx=10, pady=5)

        # 轉換按鈕
        self.convert_btn = ttk.Button(control_frame, text="開始轉換", command=self.start_conversion)
        self.convert_btn.pack(side="left", padx=5)

        # 取消按鈕
        self.cancel_btn = ttk.Button(control_frame, text="取消", command=self.cancel_conversion, state="disabled")
        self.cancel_btn.pack(side="left", padx=5)

        # 進度顯示
        progress_frame = ttk.LabelFrame(self.window, text="處理進度", padding=10)
        progress_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # 新增捲軸
        scrollbar = ttk.Scrollbar(progress_frame)
        scrollbar.pack(side="right", fill="y")

        self.progress_text = tk.Text(progress_frame, height=15, wrap="word", yscrollcommand=scrollbar.set)
        self.progress_text.pack(fill="both", expand=True)
        scrollbar.config(command=self.progress_text.yview)

        # 進度條
        self.progress_bar = ttk.Progressbar(self.window, mode='indeterminate')
        self.progress_bar.pack(fill="x", padx=10, pady=5)

        # 初始化介面狀態
        self.toggle_mode()

    def toggle_mode(self):
        """切換 YouTube/音訊 模式"""
        if self.selected_mode.get() == "youtube":
            self.youtube_frame.pack(fill="x", padx=10, pady=5)
            self.mp4_frame.pack_forget()
        else:
            self.youtube_frame.pack_forget()
            self.mp4_frame.pack(fill="x", padx=10, pady=5)

    def browse_mp4(self):
        """選擇音訊檔案"""
        try:
            filename = filedialog.askopenfilename(
                title="選擇音訊檔案",
                filetypes=[
                    ("音訊檔案", "*.m4a *.mp4 *.mp3 *.wav"),
                    ("m4a 檔案", "*.m4a"),
                    ("MP4 檔案", "*.mp4"),
                    ("MP3 檔案", "*.mp3"),
                    ("WAV 檔案", "*.wav"),
                    ("所有檔案", "*.*")
                ],
                initialdir=os.path.expanduser("~/Downloads")
            )

            if filename:
                # 檢查檔案類型
                file_ext = os.path.splitext(filename)[1].lower()
                if file_ext not in ['.m4a', '.mp4', '.mp3', '.wav']:
                    messagebox.showwarning("警告", "請選擇支援的音訊檔案格式 (m4a, mp4, mp3, wav)")
                    return

                # 使用正確的編碼顯示檔案路徑
                try:
                    display_path = filename.encode('utf-8').decode('utf-8')
                    self.mp4_path.set(display_path)
                    self.log_progress(f"已選擇檔案: {os.path.basename(display_path)}")
                except UnicodeError:
                    self.mp4_path.set(filename)
                    self.log_progress(f"已選擇檔案: {os.path.basename(filename)}")

        except Exception as e:
            messagebox.showerror("錯誤", f"選擇檔案時發生錯誤: {str(e)}")

    def log_progress(self, message):
        """更新進度顯示"""
        try:
            # 確保在主線程中更新 UI
            def update():
                self.progress_text.insert("end", message + "\n")
                self.progress_text.see("end")
                self.window.update()

            if threading.current_thread() is threading.main_thread():
                update()
            else:
                self.window.after(0, update)
        except Exception as e:
            print(f"更新進度失敗: {str(e)}")

    def update_progress_time(self):
        """更新處理時間"""
        start_time = time.time()
        while self.is_processing:
            try:
                elapsed = time.time() - start_time
                self.log_progress(f"已處理時間: {int(elapsed / 60)}分{int(elapsed % 60)}秒")
                time.sleep(10)  # 每10秒更新一次
            except:
                break

    def process_audio_chunk(self, model, audio_chunk, chunk_index, total_chunks):
        """處理音訊片段"""
        try:
            # 儲存臨時檔案
            temp_path = f"temp_chunk_{chunk_index}.wav"
            audio_chunk.export(temp_path, format="wav")

            # 處理這一段
            self.log_progress(f"\n處理第 {chunk_index + 1}/{total_chunks} 段...")
            result = model.transcribe(
                temp_path,
                language="zh",
                initial_prompt="以下是中文語音轉文字",
                condition_on_previous_text=False
            )

            # 刪除暫存檔
            os.remove(temp_path)

            return result["text"]

        except Exception as e:
            self.log_progress(f"處理片段 {chunk_index + 1} 失敗: {str(e)}")
            return ""

    def process_mp4(self):
        """處理音訊檔案"""
        if not self.is_processing:
            return

        input_path = self.mp4_path.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("錯誤", "請選擇有效的音訊檔案")
            self.finish_processing()
            return

        try:
            # 清理記憶體
            gc.collect()

            # 設定輸出路徑
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(desktop_path, f"{base_name}_transcript.txt")

            # 顯示檔案資訊
            file_size = os.path.getsize(input_path) / (1024 * 1024)  # 轉換為 MB
            self.log_progress(f"=== 處理資訊 ===")
            self.log_progress(f"檔案名稱: {base_name}")
            self.log_progress(f"檔案大小: {file_size:.2f} MB")
            self.log_progress(f"預估時間: 約 {file_size * 0.5:.1f} 分鐘")
            self.log_progress("================")

            # 載入模型
            self.log_progress("\n正在載入 Whisper 模型...")
            model = whisper.load_model("tiny")
            self.log_progress("模型載入完成")

            # 啟動進度更新線程
            progress_thread = threading.Thread(target=self.update_progress_time)
            progress_thread.daemon = True
            progress_thread.start()

            # 開始處理
            self.log_progress("\n開始轉錄音訊...")
            # 檢查檔案大小決定是否分段處理
            if file_size > 25:  # 如果大於 25MB，進行分段處理
                self.log_progress("檔案較大，將進行分段處理...")
                audio = pydub.AudioSegment.from_file(input_path)
                chunk_duration = 180 * 1000  # 3分鐘，轉換為毫秒
                chunks = []
                total_chunks = len(audio) // chunk_duration + 1

                for i in range(0, len(audio), chunk_duration):
                    if not self.is_processing:
                        break
                    chunk = audio[i:i + chunk_duration]
                    text = self.process_audio_chunk(model, chunk, i // chunk_duration, total_chunks)
                    chunks.append(text)

                result_text = " ".join(chunks)
            else:
                # 直接處理
                self.log_progress("開始轉錄（請耐心等待）...")
                result = model.transcribe(
                    input_path,
                    language="zh",
                    initial_prompt="以下是中文語音轉文字",
                    condition_on_previous_text=False
                )
                result_text = result["text"]

            # 儲存結果
            self.log_progress("\n轉錄完成，正在儲存文字檔...")
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(result_text)

            # 清理記憶體
            del model
            gc.collect()

            elapsed_time = time.time() - self.start_time if self.start_time else 0
            self.log_progress(f"\n=== 處理完成 ===")
            self.log_progress(f"總共耗時: {int(elapsed_time / 60)}分{int(elapsed_time % 60)}秒")
            self.log_progress(f"文字檔位置: {output_path}")
            messagebox.showinfo("完成", f"處理完成！\n文字稿已儲存至: {output_path}")

        except Exception as e:
            self.log_progress(f"\n處理失敗: {str(e)}")
            messagebox.showerror("錯誤", f"處理失敗: {str(e)}")

        finally:
            self.finish_processing()

    def cancel_conversion(self):
        """取消轉換程序"""
        if self.is_processing:
            self.is_processing = False
            self.log_progress("\n=== 取消處理 ===")
            self.cancel_btn.config(state="disabled")
            self.convert_btn.config(state="normal")
            self.progress_bar.stop()

    def finish_processing(self):
        """完成處理後的清理工作"""
        self.is_processing = False
        self.start_time = None
        self.progress_bar.stop()
        self.convert_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")

    def start_conversion(self):
        """開始轉換程序"""
        self.progress_text.delete(1.0, tk.END)  # 清空進度顯示
        self.log_progress("=== 開始處理 ===")

        if self.selected_mode.get() == "youtube":
            self.current_thread = threading.Thread(target=self.process_video)
        else:
            self.current_thread = threading.Thread(target=self.process_mp4)

        self.current_thread.daemon = True
        self.start_time = time.time()
        self.current_thread.start()
        self.is_processing = True
        self.cancel_btn.config(state="normal")
        self.convert_btn.config(state="disabled")
        self.progress_bar.start()

    def on_closing(self):
        """處理視窗關閉事件"""
        if self.is_processing:
            if messagebox.askokcancel("確認", "正在處理中，確定要關閉嗎？"):
                self.is_processing = False
                self.window.destroy()
        else:
            self.window.destroy()

    def run(self):
        """執行應用程式"""
        try:
            self.window.mainloop()
        except KeyboardInterrupt:
            self.on_closing()
        except Exception as e:
            print(f"發生錯誤: {str(e)}")
            self.window.destroy()


if __name__ == "__main__":
    app = VideoTranscriber()
    app.run()
import yt_dlp
import whisper
import os
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import threading
import warnings
import signal
import sys

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
        self.window.title("影片轉文字工具")
        self.window.geometry("600x500")

        # 初始化變數
        self.selected_mode = tk.StringVar(value="youtube")
        self.mp4_path = tk.StringVar()
        self.url_var = tk.StringVar()
        self.current_thread = None
        self.is_processing = False

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
        ttk.Radiobutton(mode_frame, text="本地 MP4 檔案", value="mp4",
                        variable=self.selected_mode, command=self.toggle_mode).pack(side="left", padx=5)

        # YouTube 框架
        self.youtube_frame = ttk.LabelFrame(self.window, text="YouTube 網址", padding=10)
        self.youtube_frame.pack(fill="x", padx=10, pady=5)

        self.url_entry = ttk.Entry(self.youtube_frame, textvariable=self.url_var, width=50)
        self.url_entry.pack(side="left", padx=5)

        # MP4 框架
        self.mp4_frame = ttk.LabelFrame(self.window, text="MP4 檔案", padding=10)
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

        self.progress_text = tk.Text(progress_frame, height=15, wrap="word")
        self.progress_text.pack(fill="both", expand=True)

        # 進度條
        self.progress_bar = ttk.Progressbar(self.window, mode='indeterminate')
        self.progress_bar.pack(fill="x", padx=10, pady=5)

        # 初始化介面狀態
        self.toggle_mode()

    def toggle_mode(self):
        """切換 YouTube/MP4 模式"""
        if self.selected_mode.get() == "youtube":
            self.youtube_frame.pack(fill="x", padx=10, pady=5)
            self.mp4_frame.pack_forget()
        else:
            self.youtube_frame.pack_forget()
            self.mp4_frame.pack(fill="x", padx=10, pady=5)

    def browse_mp4(self):
        """選擇 MP4 檔案"""
        filename = filedialog.askopenfilename(
            title="選擇 MP4 檔案",
            filetypes=[("MP4 檔案", "*.mp4"), ("所有檔案", "*.*")]
        )
        if filename:
            self.mp4_path.set(filename)

    def log_progress(self, message):
        """更新進度顯示"""
        try:
            self.progress_text.insert("end", message + "\n")
            self.progress_text.see("end")
            self.window.update()
        except tk.TkError:
            pass

    def _download_hook(self, d):
        """下載進度回調"""
        if d['status'] == 'downloading':
            try:
                percent = d.get('_percent_str', 'N/A')
                speed = d.get('_speed_str', 'N/A')
                self.log_progress(f"下載進度: {percent} 速度: {speed}")
            except:
                pass
        elif d['status'] == 'finished':
            self.log_progress("下載完成，開始轉換格式...")
        elif d['status'] == 'error':
            self.log_progress(f"下載錯誤: {str(d.get('error', 'unknown error'))}")

    def process_video(self):
        """處理 YouTube 影片"""
        self.is_processing = True
        url = self.url_var.get().strip()

        if not url:
            messagebox.showerror("錯誤", "請輸入 YouTube 網址")
            self.finish_processing()
            return

        if not ("youtube.com" in url or "youtu.be" in url):
            messagebox.showerror("錯誤", "請輸入有效的 YouTube 網址")
            self.finish_processing()
            return

        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            audio_path = os.path.join(desktop_path, "youtube_audio.mp3")
            text_path = os.path.join(desktop_path, "youtube_transcript.txt")

            # 下載 YouTube 音訊
            ydl_opts = {
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': '192',
                }],
                'outtmpl': audio_path.replace('.mp3', '.%(ext)s'),
                'ffmpeg_location': self.FFMPEG_PATH,
                'progress_hooks': [self._download_hook]
            }

            self.log_progress("開始下載 YouTube 音訊...")
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])

            # 轉錄文字
            self.log_progress("載入 Whisper 模型...")
            model = whisper.load_model("base", device="cpu")

            self.log_progress("開始轉錄音訊...")
            result = model.transcribe(audio_path, language="zh")

            # 儲存結果
            with open(text_path, "w", encoding="utf-8") as f:
                f.write(result["text"])

            # 清理音訊檔案
            if os.path.exists(audio_path):
                os.remove(audio_path)

            self.log_progress(f"轉錄完成！文字檔已儲存至: {text_path}")
            messagebox.showinfo("完成", f"處理完成！\n文字稿已儲存至: {text_path}")

        except Exception as e:
            self.log_progress(f"處理失敗: {str(e)}")
            messagebox.showerror("錯誤", f"處理失敗: {str(e)}")

        finally:
            self.finish_processing()

    def process_mp4(self):
        """處理 MP4 檔案"""
        if not self.is_processing:
            return

        input_path = self.mp4_path.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("錯誤", "請選擇有效的 MP4 檔案")
            self.finish_processing()
            return

        try:
            # 設定輸出路徑
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            output_path = os.path.join(desktop_path, "video_transcript.txt")

            # 直接使用 Whisper 轉換
            self.log_progress("載入 Whisper 模型...")
            model = whisper.load_model("base", device="cpu")

            self.log_progress("開始轉錄音訊...")
            result = model.transcribe(input_path, language="zh")

            # 儲存結果
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(result["text"])

            self.log_progress(f"轉錄完成！文字檔已儲存至: {output_path}")
            messagebox.showinfo("完成", f"處理完成！\n文字稿已儲存至: {output_path}")

        except Exception as e:
            self.log_progress(f"處理失敗: {str(e)}")
            messagebox.showerror("錯誤", f"處理失敗: {str(e)}")

        finally:
            self.finish_processing()

    def cancel_conversion(self):
        """取消轉換程序"""
        if self.is_processing:
            self.is_processing = False
            self.log_progress("取消處理...")
            self.cancel_btn.config(state="disabled")
            self.convert_btn.config(state="normal")
            self.progress_bar.stop()

    def finish_processing(self):
        """完成處理後的清理工作"""
        self.is_processing = False
        self.progress_bar.stop()
        self.convert_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")

    def start_conversion(self):
        """開始轉換程序"""
        if self.selected_mode.get() == "youtube":
            self.current_thread = threading.Thread(target=self.process_video)
        else:
            self.current_thread = threading.Thread(target=self.process_mp4)

        self.current_thread.daemon = True
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
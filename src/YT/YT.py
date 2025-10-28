import yt_dlp
import whisper
import os
import tkinter as tk
from tkinter import messagebox, ttk
import threading


class YouTubeTranscriber:
    def __init__(self):
        # 設定 FFmpeg 路徑
        self.FFMPEG_PATH = r"C:\Program Files\ffmpeg\bin"
        os.environ["PATH"] += os.pathsep + self.FFMPEG_PATH

        # 建立主視窗
        self.window = tk.Tk()
        self.window.title("YouTube 影片轉文字")
        self.window.geometry("600x400")

        # 建立介面
        self.create_widgets()

    def create_widgets(self):
        # YouTube 網址輸入
        url_frame = ttk.LabelFrame(self.window, text="輸入 YouTube 網址", padding=10)
        url_frame.pack(fill="x", padx=10, pady=5)

        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(url_frame, textvariable=self.url_var, width=50)
        self.url_entry.pack(side="left", padx=5)

        # 轉換按鈕
        self.convert_btn = ttk.Button(url_frame, text="開始轉換", command=self.start_conversion)
        self.convert_btn.pack(side="left", padx=5)

        # 進度顯示
        progress_frame = ttk.LabelFrame(self.window, text="處理進度", padding=10)
        progress_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.progress_text = tk.Text(progress_frame, height=15, wrap="word")
        self.progress_text.pack(fill="both", expand=True)

        # 進度條
        self.progress_bar = ttk.Progressbar(self.window, mode='indeterminate')
        self.progress_bar.pack(fill="x", padx=10, pady=5)

    def log_progress(self, message):
        """更新進度顯示"""
        self.progress_text.insert("end", message + "\n")
        self.progress_text.see("end")

    def download_audio(self, url, output_path):
        """下載 YouTube 影片的音訊"""
        ydl_opts = {
            'format': 'bestaudio/best',
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '192',
            }],
            'outtmpl': output_path,
            'ffmpeg_location': self.FFMPEG_PATH,
        }

        # 檢查 FFmpeg
        ffmpeg_exe = os.path.join(self.FFMPEG_PATH, "ffmpeg.exe")
        ffprobe_exe = os.path.join(self.FFMPEG_PATH, "ffprobe.exe")

        if not os.path.exists(ffmpeg_exe) or not os.path.exists(ffprobe_exe):
            self.log_progress(f"錯誤: 找不到 FFmpeg，請確認安裝路徑: {self.FFMPEG_PATH}")
            return False

        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])
            return True
        except Exception as e:
            self.log_progress(f"下載時發生錯誤: {str(e)}")
            return False

    def transcribe_audio(self, audio_path, output_path):
        """將音訊轉換為文字"""
        try:
            self.log_progress("載入 Whisper 模型...")
            model = whisper.load_model("base")

            self.log_progress("開始轉錄音訊...")
            result = model.transcribe(audio_path)

            with open(output_path, "w", encoding="utf-8") as f:
                f.write(result["text"])

            return True
        except Exception as e:
            self.log_progress(f"轉錄時發生錯誤: {str(e)}")
            return False

    def process_video(self):
        """處理影片的主要流程"""
        url = self.url_var.get().strip()
        if not url:
            messagebox.showerror("錯誤", "請輸入 YouTube 網址")
            return

        self.convert_btn.config(state="disabled")
        self.progress_bar.start()

        desktop_path = os.path.expanduser("~/Desktop")
        audio_path = os.path.join(desktop_path, "youtube_audio.mp3")
        text_path = os.path.join(desktop_path, "youtube_transcript.txt")

        self.log_progress("開始處理...")
        self.log_progress(f"處理網址: {url}")

        # 下載音訊
        self.log_progress("下載 YouTube 音訊...")
        if not self.download_audio(url, audio_path):
            self.log_progress("下載失敗！")
            self.finish_processing()
            return

        # 轉錄文字
        if not self.transcribe_audio(audio_path, text_path):
            self.log_progress("轉錄失敗！")
            self.finish_processing()
            return

        # 清理檔案
        try:
            os.remove(audio_path)
            self.log_progress("清理暫存檔案...")
        except:
            self.log_progress("無法刪除暫存音訊檔案")

        self.log_progress(f"\n處理完成！")
        self.log_progress(f"文字稿已儲存至: {text_path}")

        self.finish_processing()
        messagebox.showinfo("完成", f"處理完成！\n文字稿已儲存至: {text_path}")

    def finish_processing(self):
        """完成處理後的清理工作"""
        self.progress_bar.stop()
        self.convert_btn.config(state="normal")

    def start_conversion(self):
        """在新執行緒中開始轉換程序"""
        thread = threading.Thread(target=self.process_video)
        thread.daemon = True
        thread.start()

    def run(self):
        """執行應用程式"""
        self.window.mainloop()


if __name__ == "__main__":
    app = YouTubeTranscriber()
    app.run()
import yt_dlp
import whisper
import os
import tkinter as tk
from tkinter import messagebox, ttk
import threading
import warnings
import signal
import sys
import jieba
import re

# 忽略特定警告
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)


class YouTubeTranscriber:
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
        self.window.title("YouTube 影片轉文字")
        self.window.geometry("600x400")

        # 設定關閉視窗的處理
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 初始化執行緒變數
        self.current_thread = None
        self.is_processing = False

        # 建立介面
        self.create_widgets()

    def on_closing(self):
        """處理視窗關閉事件"""
        if self.is_processing:
            if messagebox.askokcancel("確認", "正在處理中，確定要關閉嗎？"):
                self.is_processing = False
                self.window.destroy()
        else:
            self.window.destroy()

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

        # 取消按鈕
        self.cancel_btn = ttk.Button(url_frame, text="取消", command=self.cancel_conversion, state="disabled")
        self.cancel_btn.pack(side="left", padx=5)

        # 進度顯示
        progress_frame = ttk.LabelFrame(self.window, text="處理進度", padding=10)
        progress_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.progress_text = tk.Text(progress_frame, height=15, wrap="word")
        self.progress_text.pack(fill="both", expand=True)

        # 進度條
        self.progress_bar = ttk.Progressbar(self.window, mode='indeterminate')
        self.progress_bar.pack(fill="x", padx=10, pady=5)

    def cancel_conversion(self):
        """取消轉換程序"""
        if self.is_processing:
            self.is_processing = False
            self.log_progress("取消處理...")
            self.cancel_btn.config(state="disabled")
            self.convert_btn.config(state="normal")
            self.progress_bar.stop()

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

    def download_audio(self, url, output_path):
        """下載 YouTube 影片的音訊"""
        if not self.is_processing:
            return False

        try:
            # 移除可能存在的舊檔案
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                    self.log_progress("移除舊的音訊檔案")
                except Exception as e:
                    self.log_progress(f"無法移除舊檔案: {str(e)}")

            # 確保輸出目錄存在
            output_dir = os.path.dirname(output_path)
            os.makedirs(output_dir, exist_ok=True)

            # 修改下載選項
            ydl_opts = {
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': '192',
                }],
                'outtmpl': output_path.replace('.mp3', '.%(ext)s'),
                'ffmpeg_location': self.FFMPEG_PATH.replace('\\', '/'),
                'verbose': True,
                'progress_hooks': [self._download_hook]
            }

            self.log_progress(f"開始下載影片...")
            self.log_progress(f"目標路徑: {output_path}")
            self.log_progress(f"FFmpeg 路徑: {self.FFMPEG_PATH}")

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                try:
                    # 先獲取影片資訊
                    info = ydl.extract_info(url, download=False)
                    title = info.get('title', 'unknown')
                    self.log_progress(f"找到影片: {title}")

                    # 開始下載
                    self.log_progress("開始下載...")
                    ydl.download([url])

                except Exception as e:
                    self.log_progress(f"YouTube-DL 錯誤: {str(e)}")
                    return False

            # 檢查檔案是否成功下載
            expected_path = output_path
            if os.path.exists(expected_path):
                file_size = os.path.getsize(expected_path)
                self.log_progress(f"下載成功! 檔案大小: {file_size / 1024 / 1024:.2f} MB")
                return True
            else:
                self.log_progress(f"找不到下載的檔案: {expected_path}")
                # 檢查目錄中的檔案
                dir_files = os.listdir(output_dir)
                self.log_progress(f"目錄中的檔案: {dir_files}")
                return False

        except Exception as e:
            self.log_progress(f"下載過程發生錯誤: {str(e)}")
            import traceback
            self.log_progress(traceback.format_exc())
            return False

    def transcribe_audio(self, audio_path, output_path):
        """將音訊轉換為文字"""
        if not self.is_processing:
            return False

        try:
            # 檢查音訊檔案是否存在
            if not os.path.exists(audio_path):
                self.log_progress(f"錯誤：找不到音訊檔案 {audio_path}")
                return False

            self.log_progress("載入 Whisper 模型...")
            try:
                model = whisper.load_model("base", device="cpu")
            except Exception as e:
                self.log_progress(f"載入模型失敗: {str(e)}")
                return False

            self.log_progress("開始轉錄音訊...")
            try:
                audio_path_str = str(audio_path)
                options = {
                    "language": "zh",
                    "task": "transcribe",
                    "fp16": False
                }
                result = model.transcribe(audio_path_str, **options)
            except Exception as e:
                self.log_progress(f"轉錄過程失敗: {str(e)}")
                return False

            # 處理文字並加入標點符號
            try:
                segments = result["segments"]
                formatted_text = ""

                for segment in segments:
                    # 取得該段文字並清理
                    segment_text = segment["text"].strip()
                    if not segment_text:
                        continue

                    # 使用結巴分詞
                    words = list(jieba.cut(segment_text))

                    # 根據分詞結果重新組織句子
                    current_sentence = ""
                    for word in words:
                        current_sentence += word

                        # 如果詞語本身就是標點符號，直接加入
                        if re.match(r'[，。！？、：；]', word):
                            formatted_text += current_sentence
                            current_sentence = ""

                    # 處理剩餘的句子
                    if current_sentence:
                        # 如果句子結尾沒有標點符號，加入句號
                        if not re.search(r'[。！？]$', current_sentence):
                            current_sentence += "。"
                        formatted_text += current_sentence

                    # 每個段落後添加換行
                    formatted_text += "\n"

                # 最終清理：移除多餘的空行和空格
                formatted_text = re.sub(r'\n+', '\n', formatted_text.strip())

                # 寫入結果
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(formatted_text)
                self.log_progress(f"成功將轉錄結果寫入: {output_path}")

            except Exception as e:
                self.log_progress(f"寫入檔案失敗: {str(e)}")
                return False

            return True
        except Exception as e:
            self.log_progress(f"轉錄時發生未預期的錯誤: {str(e)}")
            return False

    def process_video(self):
        """處理影片的主要流程"""
        self.is_processing = True
        self.cancel_btn.config(state="normal")

        url = self.url_var.get().strip()
        if not url:
            messagebox.showerror("錯誤", "請輸入 YouTube 網址")
            self.finish_processing()
            return

        # 驗證 URL 格式
        if not ("youtube.com" in url or "youtu.be" in url):
            messagebox.showerror("錯誤", "請輸入有效的 YouTube 網址")
            self.finish_processing()
            return

        self.convert_btn.config(state="disabled")
        self.progress_bar.start()

        try:
            # 使用絕對路徑
            desktop_path = os.path.abspath(os.path.expanduser("~/Desktop"))
            audio_path = os.path.join(desktop_path, "youtube_audio.mp3")
            text_path = os.path.join(desktop_path, "youtube_transcript.txt")

            # 轉換為正斜線
            audio_path = audio_path.replace('\\', '/')
            text_path = text_path.replace('\\', '/')

            self.log_progress(f"桌面路徑: {desktop_path}")
            self.log_progress(f"音訊檔案路徑: {audio_path}")
            self.log_progress(f"文字檔案路徑: {text_path}")

            # 檢查路徑是否可寫入
            try:
                with open(audio_path, 'w') as f:
                    pass
                os.remove(audio_path)
                self.log_progress("檔案路徑可寫入")
            except Exception as e:
                self.log_progress(f"檔案路徑無法寫入: {str(e)}")
                raise Exception("無法寫入檔案，請檢查權限")

            if not self.is_processing:
                return

            self.log_progress("開始處理...")
            self.log_progress(f"處理網址: {url}")

            # 下載音訊
            self.log_progress("下載 YouTube 音訊...")
            if not self.download_audio(url, audio_path):
                raise Exception("下載失敗")

            # 確認音訊檔案存在
            if not os.path.exists(audio_path):
                raise Exception(f"找不到下載的音訊檔案: {audio_path}")

            if not self.is_processing:
                return

            # 轉錄文字
            self.log_progress("開始轉錄音訊為文字...")
            if not self.transcribe_audio(audio_path, text_path):
                raise Exception("轉錄失敗")

            # 確認文字檔案已經生成
            if not os.path.exists(text_path):
                raise Exception("轉錄完成但找不到輸出的文字檔")

            # 清理檔案
            try:
                if os.path.exists(audio_path):
                    os.remove(audio_path)
                    self.log_progress("清理暫存檔案...")
            except Exception as e:
                self.log_progress(f"無法刪除暫存音訊檔案: {str(e)}")

            if self.is_processing:
                self.log_progress(f"\n處理完成！")
                self.log_progress(f"文字稿已儲存至: {text_path}")
                messagebox.showinfo("完成", f"處理完成！\n文字稿已儲存至: {text_path}")

        except Exception as e:
            if self.is_processing:
                error_msg = f"處理失敗: {str(e)}"
                self.log_progress(error_msg)
                messagebox.showerror("錯誤", error_msg)

        finally:
            self.finish_processing()

    def finish_processing(self):
        """完成處理後的清理工作"""
        self.is_processing = False
        self.progress_bar.stop()
        self.convert_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")

    def start_conversion(self):
        """在新執行緒中開始轉換程序"""
        self.current_thread = threading.Thread(target=self.process_video)
        self.current_thread.daemon = True
        self.current_thread.start()

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
    app = YouTubeTranscriber()
    app.run()
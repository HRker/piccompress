import yt_dlp
import whisper
from transformers import pipeline


def download_audio(url, output_file="audio.mp3"):
    """下載YouTube影片的音訊"""
    ydl_opts = {
        'format': 'bestaudio/best',
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }],
        'outtmpl': output_file,
    }

    with yt_dlp.YDL(ydl_opts) as ydl:
        ydl.download([url])


def transcribe_audio(audio_file):
    """使用Whisper將音訊轉換為文字"""
    model = whisper.load_model("base")
    result = model.transcribe(audio_file)
    return result["text"]


def translate_text(text):
    """將英文文本翻譯成中文"""
    translator = pipeline("translation", model="Helsinki-NLP/opus-mt-en-zh")
    translated = translator(text)
    return translated[0]['translation_text']


def process_youtube_video(url):
    """主要處理函數"""
    # 1. 下載音訊
    print("正在下載音訊...")
    download_audio(url)

    # 2. 轉錄成文字
    print("正在轉錄音訊...")
    transcript = transcribe_audio("audio.mp3")

    # 3. 翻譯成中文
    print("正在翻譯...")
    chinese_text = translate_text(transcript)

    return transcript, chinese_text


# 使用示例
if __name__ == "__main__":
    youtube_url = "你的YouTube網址"
    english_text, chinese_text = process_youtube_video(youtube_url)

    print("\n英文文字稿:")
    print(english_text)
    print("\n中文翻譯:")
    print(chinese_text)
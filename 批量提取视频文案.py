import os
import subprocess
import requests
import openpyxl
import hashlib  # 用于计算文件的哈希值

# 请替换为您的 SiliconFlow API 密钥
SILICONFLOW_API_KEY = "sk-nexfkxivirurtdvbpzkjgjcyplorwkssfitcvnppeaclunbe"
# 模型选择：FunAudioLLM 或 SenseVoiceSmall。
MODEL_NAME = "FunAudioLLM/SenseVoiceSmall"
# 缓存文件路径
CACHE_FILE = "transcription_cache.txt"


def load_cache():
    """加载缓存，从文件中读取已处理过的音频文件的哈希值。"""
    cache = set()
    try:
        with open(CACHE_FILE, "r") as f:
            for line in f:
                cache.add(line.strip())
    except FileNotFoundError:
        pass  # 缓存文件不存在，正常
    return cache


def save_cache(cache):
    """保存缓存，将已处理过的音频文件的哈希值写入文件。"""
    with open(CACHE_FILE, "w") as f:
        for hash_value in cache:
            f.write(hash_value + "\n")


def get_file_hash(filepath):
    """计算文件的 SHA256 哈希值。"""
    hasher = hashlib.sha256()
    with open(filepath, "rb") as f:
        while True:
            chunk = f.read(4096)  # 分块读取，避免大文件占用过多内存
            if not chunk:
                break
            hasher.update(chunk)
    return hasher.hexdigest()


def extract_audio(video_path, audio_path):
    """
    使用 ffmpeg 从视频文件中提取音频。  (与之前相同)
    """
    try:
        command = [
            "ffmpeg",
            "-i", video_path,
            "-vn",
            "-acodec", "pcm_s16le",
            "-ar", "16000",
            "-ac", "1",
            audio_path
        ]
        subprocess.run(command, check=True, capture_output=True)
        print(f"音频提取成功: {audio_path}")
    except subprocess.CalledProcessError as e:
        print(f"音频提取失败: {video_path}")
        print(e.stderr.decode())


def transcribe_audio(audio_path):
    """
    使用 SiliconFlow API 将音频转录为文本。(与之前相同，使用正确的 multipart/form-data 请求)
    """
    url = "https://api.siliconflow.cn/v1/audio/transcriptions"
    headers = {
        "Authorization": f"Bearer {SILICONFLOW_API_KEY}",
    }

    files = {
        "file": (os.path.basename(audio_path), open(audio_path, "rb")),
        "model": (None, MODEL_NAME),
    }

    try:
        response = requests.post(url, headers=headers, files=files)
        response.raise_for_status()
        return response.json()["text"]
    except requests.exceptions.RequestException as e:
        print(f"音频转录失败: {audio_path}")
        print(e)
        return None


def write_to_excel(text_lines, excel_path="1.xlsx"):
    """
    将文本逐行写入 Excel 文件。(与之前相同)
    """
    try:
        if not os.path.exists(excel_path):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
        else:
            workbook = openpyxl.load_workbook(excel_path)
            sheet = workbook.active

        for line in text_lines:
            sheet.append([line])

        workbook.save(excel_path)
        print(f"文本已写入: {excel_path}")

    except Exception as e:
        print(f"写入 Excel 失败: {e}")


def main():
    """
    主函数，遍历视频文件，提取音频，转录文本，写入 Excel。
    """
    cache = load_cache()  # 加载缓存

    for filename in os.listdir("."):
        if filename.endswith((".mp4", ".avi", ".mov", ".mkv")):
            video_path = os.path.join(".", filename)
            print(f"处理视频: {video_path}")

            audio_filename = os.path.splitext(filename)[0] + ".wav"
            audio_path = os.path.join(".", audio_filename)

            # 计算视频文件的哈希值
            video_hash = get_file_hash(video_path)

            # 检查视频文件是否已处理过
            if video_hash in cache:
                print(f"跳过已处理的视频: {video_path}")
                continue

            extract_audio(video_path, audio_path)
            transcribed_text = transcribe_audio(audio_path)

            if transcribed_text:
                text_lines = transcribed_text.splitlines()
                write_to_excel(text_lines)

                # 将视频文件的哈希值添加到缓存
                cache.add(video_hash)
                save_cache(cache)  # 保存缓存

            # (可选) 删除提取的音频文件
            # os.remove(audio_path)


if __name__ == "__main__":
    main()

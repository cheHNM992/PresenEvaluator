# プレゼン評価システム

import os
import sys
import re
import pptx
import openai
import numpy as np
import whisper
from pptx import Presentation
from datetime import datetime

# ==== 設定 ====
os.environ['OPENAI_API_KEY'] = 'sk-proj-*****'

# ==== 音声分析モジュール ====
def transcribe_audio(file_path):
    model = whisper.load_model("base")
    result = model.transcribe(file_path, fp16=False, language='ja')

    # 出力ファイル名にタイムスタンプを付加
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"transcription_{timestamp}.txt"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(result['text'])

    return result['text'], result['segments']


def analyze_speech(segments):
    total_words = sum(len(seg['text'].split()) for seg in segments)
    duration_minutes = (segments[-1]['end'] - segments[0]['start']) / 60.0
    wpm = total_words / duration_minutes if duration_minutes else 0

    filler_words = ['えーと', 'あの', 'えっと', 'その']
    filler_count = sum(sum(word in seg['text'] for word in filler_words) for seg in segments)

    pause_lengths = [segments[i+1]['start'] - segments[i]['end'] for i in range(len(segments)-1)]
    long_pauses = sum(1 for p in pause_lengths if p > 1.0)

    return {
        "wpm": round(wpm, 2),
        "filler_count": filler_count,
        "long_pauses": long_pauses
    }


# ==== 資料分析モジュール ====
def extract_ppt_text(file_path):
    prs = Presentation(file_path)
    slides_info = []
    for i, slide in enumerate(prs.slides):
        slide_text = "\n".join([shape.text for shape in slide.shapes if hasattr(shape, "text")])
        slide_chars = len(slide_text)
        image_count = sum(1 for shape in slide.shapes if "Picture" in shape.name)
        slides_info.append({
            "slide_number": i + 1,
            "text": slide_text,
            "char_count": slide_chars,
            "image_count": image_count
        })
    return slides_info


def analyze_slide_structure(slides_info):
    avg_chars = np.mean([s['char_count'] for s in slides_info])
    total_images = sum(s['image_count'] for s in slides_info)
    return {
        "avg_chars_per_slide": round(avg_chars, 1),
        "total_images": total_images,
        "slide_count": len(slides_info)
    }


# ==== 評価とアドバイス生成 ====
def generate_evaluation(transcription, slide_data):
    from openai import OpenAI
    client = OpenAI()

    prompt = f"""
以下はプレゼンの文字起こしとスライド内容の要約です。

[文字起こし]:
{transcription}

[スライド概要]:
{slide_data}

以下の4つの観点（内容、プレゼン技術、視覚資料、構成）について、それぞれ5段階で評価し、簡単な理由と改善点、長所を出力してください。
最後に3つの改善点と具体的なアドバイスも示してください。

フォーマットは必ず以下としてください：
内容: ○点
プレゼン技術: ○点
視覚資料: ○点
構成: ○点

その後に評価コメントを書いてください。
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        messages=[
            {"role": "system", "content": "あなたはプロのプレゼン評価者です。"},
            {"role": "user", "content": prompt}
        ]
    )

    content = response.choices[0].message.content
    return content


# ==== スコア抽出 ====
def extract_scores(evaluation_text):
    pattern = r"内容: (\d)点.*?プレゼン技術: (\d)点.*?視覚資料: (\d)点.*?構成: (\d)点"
    match = re.search(pattern, evaluation_text, re.DOTALL)

    if match:
        return {
            "内容": int(match.group(1)),
            "プレゼン技術": int(match.group(2)),
            "視覚資料": int(match.group(3)),
            "構成": int(match.group(4))
        }
    else:
        print("スコアの抽出に失敗しました。デフォルトで全て0点とします。")
        return {
            "内容": 0,
            "プレゼン技術": 0,
            "視覚資料": 0,
            "構成": 0
        }


# ==== スコア集計 ====
def compute_score(sub_scores):
    weights = {
        "内容": 0.3,
        "プレゼン技術": 0.3,
        "視覚資料": 0.2,
        "構成": 0.2
    }
    total = sum(sub_scores[k] * weights[k] for k in weights)
    return round(total * 20, 1)  # 100点満点換算


# ==== プレゼン評価処理 ====
def evaluate_presentation(audio_path, ppt_path):
    text, segments = transcribe_audio(audio_path)
    speech_analysis = analyze_speech(segments)

    slides_info = extract_ppt_text(ppt_path)
    slide_analysis = analyze_slide_structure(slides_info)

    slide_summary = f"スライド数: {slide_analysis['slide_count']}, 平均文字数: {slide_analysis['avg_chars_per_slide']}, 画像数: {slide_analysis['total_images']}"

    evaluation = generate_evaluation(text, slide_summary)

    # GPTの出力からスコアを抽出
    sub_scores = extract_scores(evaluation)
    total_score = compute_score(sub_scores)

    print("==== 音声分析 ====")
    print(speech_analysis)
    print("\n==== スライド分析 ====")
    print(slide_analysis)
    print("\n==== GPT評価 ====")
    print(evaluation)
    print(f"\n==== 総合得点: {total_score}点 ====")

    # 評価結果をファイルに保存
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    result_filename = f"evaluation_result_{timestamp}.txt"

    with open(result_filename, "w", encoding="utf-8") as f:
        f.write("==== 音声分析 ====\n")
        f.write(str(speech_analysis) + "\n\n")
        f.write("==== スライド分析 ====\n")
        f.write(str(slide_analysis) + "\n\n")
        f.write("==== GPT評価 ====\n")
        f.write(evaluation + "\n\n")
        f.write(f"==== 総合得点: {total_score}点 ====\n")

    print(f"\n評価結果をファイルに保存しました: {result_filename}")


# ==== メイン関数 ====
def main():
    if len(sys.argv) != 3:
        print("使用方法: python script.py 音声ファイル パワポファイル")
        print("例: python script.py sample.wav slides.pptx")
        sys.exit(1)

    audio_path = sys.argv[1]
    ppt_path = sys.argv[2]

    if not os.path.exists(audio_path):
        print(f"音声ファイルが見つかりません: {audio_path}")
        sys.exit(1)
    if not os.path.exists(ppt_path):
        print(f"PowerPointファイルが見つかりません: {ppt_path}")
        sys.exit(1)

    evaluate_presentation(audio_path, ppt_path)


if __name__ == "__main__":
    main()

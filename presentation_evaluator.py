# プレゼン評価システム

import os
import sys
import re
import pptx
import openai
import numpy as np
import base64
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from datetime import datetime

# ==== 設定 ====
openai.api_key = 'sk-proj-*****'  # ご自身のAPIキーに置き換えてください

# ==== 音声分析モジュール ====
def transcribe_audio(file_path):
    audio_file = open(file_path, "rb")
    response = openai.audio.transcriptions.create(
        model="whisper-1",
        file=audio_file,
        response_format="verbose_json",
        language="ja"
    )

    text = response.text
    segments = response.segments

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"transcription_{timestamp}.txt"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(text)

    return text, segments


def analyze_speech(segments):
    total_words = sum(len(seg.text.split()) for seg in segments)
    duration_minutes = (segments[-1].end - segments[0].start) / 60.0
    wpm = total_words / duration_minutes if duration_minutes else 0

    filler_words = ['えーと', 'あの', 'えっと', 'その']
    filler_count = sum(sum(word in seg.text for word in filler_words) for seg in segments)

    pause_lengths = [segments[i + 1].start - segments[i].end for i in range(len(segments) - 1)]
    long_pauses = sum(1 for p in pause_lengths if p > 1.0)

    return {
        "wpm": round(wpm, 2),
        "filler_count": filler_count,
        "long_pauses": long_pauses
    }


# ==== 資料抽出モジュール ====
def extract_ppt_text(file_path):
    prs = Presentation(file_path)
    slides_text = []
    for i, slide in enumerate(prs.slides):
        slide_text = "\n".join([shape.text for shape in slide.shapes if hasattr(shape, "text")])
        slides_text.append(f"スライド {i + 1}:\n{slide_text}\n")
    return "\n".join(slides_text)


def extract_images_from_ppt(ppt_path, output_dir):
    prs = Presentation(ppt_path)
    image_files = []

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for i, slide in enumerate(prs.slides):
        for j, shape in enumerate(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image = shape.image
                image_bytes = image.blob
                image_filename = os.path.join(output_dir, f"slide_{i + 1}_image_{j + 1}.png")
                with open(image_filename, 'wb') as f:
                    f.write(image_bytes)
                image_files.append(image_filename)

    return image_files


def encode_image_to_base64(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')


def analyze_image(image_path):
    base64_image = encode_image_to_base64(image_path)
    response = openai.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "あなたは画像解析の専門家です。"},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "この画像に何が写っているか説明し、プレゼン資料として適切か評価してください。視覚資料としての質も5段階（小数点1桁まで）で採点してください。フォーマット: 視覚資料: ○.○点"},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}
                ]
            }
        ]
    )
    return response.choices[0].message.content


def extract_visual_score(image_analysis):
    pattern = r"視覚資料: ([0-5](?:\.\d)?)点"
    matches = re.findall(pattern, image_analysis)
    if matches:
        scores = [float(score) for score in matches]
        return round(sum(scores) / len(scores), 1)
    return 0.0


def analyze_all_images(image_files):
    all_analyses = []
    for image_path in image_files:
        print(f"画像を解析中: {image_path}")
        analysis = analyze_image(image_path)
        all_analyses.append(f"{image_path}:\n{analysis}\n")
    return "\n".join(all_analyses)


def analyze_slide(slide_text):
    response = openai.chat.completions.create(
        model="gpt-4.1",
        messages=[
            {"role": "system", "content": "あなたはプロのプレゼン資料評価者です。"},
            {"role": "user", "content": f"""
以下はプレゼンテーションのスライド全文です。

[スライド全文]
{slide_text}

この資料のスライド数、各スライドの文字量の適切さ、視覚的な情報量（図表や画像の有無）を評価し、全体的な視覚資料の質を以下のフォーマットで5段階（小数点1桁まで）で評価してください。

視覚資料: ○.○点

その後に資料の良い点と改善点を簡単にまとめてください。
"""}
        ]
    )
    return response.choices[0].message.content


def generate_evaluation_with_images(transcription, slide_analysis, image_analysis):
    response = openai.chat.completions.create(
        model="gpt-4.1",
        messages=[
            {"role": "system", "content": "あなたはプロのプレゼン評価者です。"},
            {"role": "user", "content": f"""
以下はプレゼンの文字起こしとスライド資料の分析結果、および画像分析結果です。

[文字起こし]:
{transcription}

[スライド分析]:
{slide_analysis}

[画像分析]:
{image_analysis}

以下の4つの観点（内容、プレゼン技術、視覚資料、構成）について、それぞれ5段階（小数点1桁まで）で評価し、簡単な理由と改善点、長所を出力してください。
最後に3つの改善点と具体的なアドバイスも示してください。

フォーマットは必ず以下としてください：
内容: ○.○点
プレゼン技術: ○.○点
視覚資料: ○.○点
構成: ○.○点

その後に評価コメントを書いてください。
"""}
        ]
    )
    return response.choices[0].message.content


# ==== スコア抽出 ====
def extract_scores(evaluation_text):
    pattern = r"内容: ([0-5](?:\.\d)?)点.*?プレゼン技術: ([0-5](?:\.\d)?)点.*?視覚資料: ([0-5](?:\.\d)?)点.*?構成: ([0-5](?:\.\d)?)点"
    match = re.search(pattern, evaluation_text, re.DOTALL)

    if match:
        return {
            "内容": float(match.group(1)),
            "プレゼン技術": float(match.group(2)),
            "視覚資料": float(match.group(3)),
            "構成": float(match.group(4))
        }
    else:
        print("スコアの抽出に失敗しました。デフォルトで全て0.0点とします。")
        return {
            "内容": 0.0,
            "プレゼン技術": 0.0,
            "視覚資料": 0.0,
            "構成": 0.0
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
    return round(total * 20, 1)


# ==== プレゼン評価処理 ====
def evaluate_presentation(audio_path, ppt_path):
    text, segments = transcribe_audio(audio_path)
    speech_analysis = analyze_speech(segments)

    slides_text = extract_ppt_text(ppt_path)
    slide_analysis = analyze_slide(slides_text)

    image_files = extract_images_from_ppt(ppt_path, "extracted_images")
    if image_files:
        image_analysis = analyze_all_images(image_files)
        image_visual_score = extract_visual_score(image_analysis)
    else:
        image_analysis = "画像は含まれていません。"
        image_visual_score = 0.0

    evaluation = generate_evaluation_with_images(text, slide_analysis, image_analysis)

    sub_scores = extract_scores(evaluation)
    sub_scores["視覚資料"] = round((sub_scores["視覚資料"] + image_visual_score) / 2, 1)

    total_score = compute_score(sub_scores)

    print("==== 音声分析 ====")
    print(speech_analysis)
    print("\n==== スライド分析 ====")
    print(slide_analysis)
    print("\n==== 画像分析 ====")
    print(image_analysis)
    print("\n==== 総合評価 ====")
    print(evaluation)
    print(f"\n==== 総合得点: {total_score}点 ====")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    result_filename = f"evaluation_result_{timestamp}.txt"

    with open(result_filename, "w", encoding="utf-8") as f:
        f.write("==== 音声分析 ====\n")
        f.write(str(speech_analysis) + "\n\n")
        f.write("==== スライド分析 ====\n")
        f.write(slide_analysis + "\n\n")
        f.write("==== 画像分析 ====\n")
        f.write(image_analysis + "\n\n")
        f.write("==== 総合評価 ====\n")
        f.write(evaluation + "\n\n")
        f.write(f"==== 総合得点: {total_score}点 ====\n")

    print(f"\n評価結果をファイルに保存しました: {result_filename}")

    # 画像ファイル自動削除
    if image_files:
        for image_path in image_files:
            os.remove(image_path)
        os.rmdir("extracted_images")
        print("\n一時画像ファイルを削除しました。")


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

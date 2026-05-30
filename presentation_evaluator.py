# プレゼン評価システム (CLI/GUI統合版)

import os
import sys
import re
import base64
import shutil
from datetime import datetime

# Streamlitがインポート可能かチェック
try:
    import streamlit as st
    STREAMLIT_AVAILABLE = True
except ImportError:
    STREAMLIT_AVAILABLE = False

import openai
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


# ==== グローバル設定 ====
MODEL_LLM = "gpt-5.4"
# MODEL_LLM = "gpt-5.4-nano"
MODEL_WHISPER = "whisper-1"


# ==== 共通関数群 ====

def transcribe_audio(file_path, client):
    """音声ファイルをテキストに変換"""
    audio_file = open(file_path, "rb")
    response = client.audio.transcriptions.create(
        model=MODEL_WHISPER,
        file=audio_file,
        response_format="verbose_json",
        language="ja"
    )

    text = response.text
    segments = response.segments

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    audio_filename = os.path.splitext(os.path.basename(file_path))[0]
    filename = f"transcription_{timestamp}_{audio_filename}.txt"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(text)

    return text, segments


def analyze_speech(segments):
    """音声分析: WPM、フィラーワード、長い間の検出"""
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


def extract_ppt_text(file_path):
    """PowerPointからテキストを抽出"""
    prs = Presentation(file_path)
    slides_text = []
    for i, slide in enumerate(prs.slides):
        slide_text = "\n".join([shape.text for shape in slide.shapes if hasattr(shape, "text")])
        slides_text.append(f"スライド {i + 1}:\n{slide_text}\n")
    return "\n".join(slides_text)


def extract_images_from_ppt(ppt_path, output_dir):
    """PowerPointから画像を抽出"""
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
    """画像をBase64エンコード"""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')


def analyze_image(image_path, client):
    """画像をAIで分析"""
    print(f"ファイル名: {image_path}")  # 画像分析で失敗する可能性があるため、デバッグ用に出力を追加
    base64_image = encode_image_to_base64(image_path)
    response = client.chat.completions.create(
        model=MODEL_LLM,
        messages=[
            {"role": "system", "content": "あなたは画像解析の専門家です。"},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "この画像に何が写っているか説明し、プレゼン資料として適切か評価してください。視覚資料としての質も100点満点(整数)で採点してください。フォーマット: 視覚資料: ○点"},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}
                ]
            }
        ]
    )
    return response.choices[0].message.content


def extract_visual_score(image_analysis):
    """画像分析結果からスコアを抽出"""
    pattern = r"視覚資料: ([0-9]{1,3})点"
    matches = re.findall(pattern, image_analysis)
    if matches:
        scores = [int(score) for score in matches]
        return int(sum(scores) / len(scores))
    return 0


def analyze_all_images(image_files, client):
    """全画像を分析"""
    all_analyses = []
    for image_path in image_files:
        analysis = analyze_image(image_path, client)
        all_analyses.append(f"{image_path}:\n{analysis}\n")
    return "\n".join(all_analyses)


def analyze_slide_text(slide_text, client):
    """スライドテキストを分析"""
    response = client.chat.completions.create(
        model=MODEL_LLM,
        messages=[
            {"role": "system", "content": "あなたはプロのプレゼン資料評価者です。"},
            {"role": "user", "content": f"""
以下はプレゼンテーションのスライド全文です。

[スライド全文]
{slide_text}

この資料のスライド数、各スライドの文字量の適切さ、内容を評価し、全体的な資料の質を以下のフォーマットで100点満点(整数)で評価してください。
ただし、図表や画像は評価に含めないでください。

資料: ○点

その後に資料の良い点と改善点を簡単にまとめてください。
"""}
        ]
    )
    return response.choices[0].message.content


def generate_evaluation_with_images(transcription, slide_text_analysis, image_analysis, client):
    """総合評価を生成"""
    response = client.chat.completions.create(
        model=MODEL_LLM,
        messages=[
            {"role": "system", "content": "あなたはプロのプレゼン評価者です。"},
            {"role": "user", "content": f"""
以下はプレゼンの文字起こしとスライド資料の分析結果、および画像分析結果です。

[文字起こし]:
{transcription}

[スライドテキスト分析]:
{slide_text_analysis}

[画像分析]:
{image_analysis}

以下の4つの観点(内容、プレゼン技術、視覚資料、構成)について、それぞれ100点満点(整数)で評価し、簡単な理由と改善点、長所を出力してください。
最後に3つの改善点と具体的なアドバイスも示してください。

フォーマットは必ず以下としてください:
内容: ○点
プレゼン技術: ○点
視覚資料: ○点
構成: ○点

その後に評価コメントを書いてください。
"""}
        ]
    )
    return response.choices[0].message.content


def extract_scores(evaluation_text):
    """評価テキストからスコアを抽出"""
    pattern = r"内容: ([0-9]{1,3})点.*?プレゼン技術: ([0-9]{1,3})点.*?視覚資料: ([0-9]{1,3})点.*?構成: ([0-9]{1,3})点"
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


def compute_score(sub_scores):
    """サブスコアから総合スコアを計算"""
    weights = {
        "内容": 0.3,
        "プレゼン技術": 0.3,
        "視覚資料": 0.2,
        "構成": 0.2
    }
    total = sum(float(sub_scores[k]) * weights[k] for k in weights)
    return int(round(total, 0))


def evaluate_presentation_core(audio_path, ppt_path, client, progress_callback=None):
    """
    プレゼン評価のコア処理
    progress_callback: 進捗を通知するコールバック関数(GUI用)
    """
    def log(message):
        print(message)
        if progress_callback:
            progress_callback(message)

    log("音声分析中")
    text, segments = transcribe_audio(audio_path, client)
    speech_analysis = analyze_speech(segments)

    log("資料分析中")
    slides_text = extract_ppt_text(ppt_path)
    slide_text_analysis = analyze_slide_text(slides_text, client)

    log("画像解析中")
    image_files = extract_images_from_ppt(ppt_path, "extracted_images")
    if image_files:
        image_analysis = analyze_all_images(image_files, client)
        image_visual_score = extract_visual_score(image_analysis)
    else:
        image_analysis = "画像は含まれていません。"
        image_visual_score = 0

    log("総合評価生成中")
    evaluation = generate_evaluation_with_images(text, slide_text_analysis, image_analysis, client)

    sub_scores = extract_scores(evaluation)
    # 画像なしの場合は画像の得点を判定しないようにする
    if image_visual_score == 0:
        image_visual_score = sub_scores["視覚資料"]
    sub_scores["視覚資料"] = int(round((sub_scores["視覚資料"] + image_visual_score) / 2, 0))

    total_score = compute_score(sub_scores)

    # 結果を保存
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    ppt_filename = os.path.splitext(os.path.basename(ppt_path))[0]
    result_filename = f"evaluation_result_{timestamp}_{ppt_filename}.txt"

    with open(result_filename, "w", encoding="utf-8") as f:
        f.write(f"==== 総合得点: {total_score}点 ====\n\n")
        f.write("==== 総合評価 ====\n")
        f.write(evaluation + "\n\n")
        f.write("==== 音声分析 ====\n")
        f.write(str(speech_analysis) + "\n\n")

    log(f"評価結果をファイルに保存しました: {result_filename}")

    # 一時画像ファイルを削除
    if image_files:
        for image_path in image_files:
            if os.path.exists(image_path):
                os.remove(image_path)
        if os.path.exists("extracted_images"):
            os.rmdir("extracted_images")
        log("一時画像ファイルを削除しました。")

    return {
        "total_score": total_score,
        "sub_scores": sub_scores,
        "evaluation": evaluation,
        "speech_analysis": speech_analysis,
        "transcription": text,
        "slide_text_analysis": slide_text_analysis,
        "image_analysis": image_analysis,
        "image_files": image_files
    }


# ==== CLIモード ====
def run_cli_mode():
    """コマンドライン実行モード"""
    if len(sys.argv) != 3:
        print("使用方法: python presentation_evaluator.py 音声ファイル パワポファイル")
        print("例: python presentation_evaluator.py sample.wav slides.pptx")
        sys.exit(1)

    audio_path = sys.argv[1]
    ppt_path = sys.argv[2]

    if not os.path.exists(audio_path):
        print(f"音声ファイルが見つかりません: {audio_path}")
        sys.exit(1)
    if not os.path.exists(ppt_path):
        print(f"PowerPointファイルが見つかりません: {ppt_path}")
        sys.exit(1)

    # APIキーは環境変数から取得
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        print("エラー: OPENAI_API_KEY環境変数が設定されていません")
        sys.exit(1)

    client = openai.OpenAI(api_key=api_key)
    
    evaluate_presentation_core(audio_path, ppt_path, client)


# ==== GUIモード ====
def run_gui_mode():
    """Streamlit GUIモード"""
    if not STREAMLIT_AVAILABLE:
        print("エラー: Streamlitがインストールされていません")
        print("インストール: pip install streamlit")
        sys.exit(1)

    # ページ設定
    st.set_page_config(page_title="AIプレゼン評価システム", layout="wide")

    # スタイル設定
    st.markdown("""
        <style>
        .main {
            background-color: #f8f9fa;
        }
        .stMetric {
            background-color: #ffffff;
            padding: 15px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        </style>
        """, unsafe_allow_html=True)

    # タイトル・説明
    st.title("🎤 AIプレゼン評価システム")
    st.markdown("音声ファイルとスライド資料をアップロードするだけで、AIがあなたのプレゼンを多角的に分析・採点します。")

    # サイドバー設定
    with st.sidebar:
        st.header("⚙️ 設定")
        api_key = st.text_input("OpenAI API Keyを入力してください", type="password")
        
        st.info(f"""
        **使用モデル:**
        - LLM: {MODEL_LLM}
        - 音声: {MODEL_WHISPER}
        
        **分析項目:**
        1. 内容 (30%)
        2. プレゼン技術 (30%)
        3. 視覚資料 (20%)
        4. 構成 (20%)
        """)

    # APIキーのチェック
    if not api_key:
        st.warning("⚠️ 続行するにはサイドバーにOpenAI APIキーを入力してください。")
        st.stop()

    client = openai.OpenAI(api_key=api_key)

    # ファイルアップロード
    col1, col2 = st.columns(2)
    with col1:
        audio_upload = st.file_uploader("1. 音声ファイルをアップロード", type=['mp3', 'wav', 'm4a', 'mp4'])
    with col2:
        ppt_upload = st.file_uploader("2. PowerPointファイルをアップロード", type=['pptx'])

    if st.button("📊 プレゼンを分析する", use_container_width=True):
        if audio_upload and ppt_upload:
            # 一時ファイル保存
            temp_dir = "temp_process"
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)
            
            audio_path = os.path.join(temp_dir, audio_upload.name)
            ppt_path = os.path.join(temp_dir, ppt_upload.name)
            
            with open(audio_path, "wb") as f:
                f.write(audio_upload.getbuffer())
            with open(ppt_path, "wb") as f:
                f.write(ppt_upload.getbuffer())

            try:
                with st.status("分析中...", expanded=True) as status:
                    progress_messages = []
                    
                    def progress_callback(msg):
                        progress_messages.append(msg)
                        icon_map = {
                            "音声分析中": "🎙️",
                            "資料分析中": "📄",
                            "画像解析中": "🖼️",
                            "総合評価生成中": "🤖"
                        }
                        icon = icon_map.get(msg, "⏳")
                        st.write(f"{icon} {msg}...")

                    result = evaluate_presentation_core(
                        audio_path, ppt_path, client, 
                        progress_callback=progress_callback
                    )

                    status.update(label="✅ 分析が完了しました！", state="complete", expanded=False)

                # 結果表示
                st.divider()
                
                tab1, tab2, tab3 = st.tabs(["📝 総合評価レポート", "📖 文字起こし全文", "🖼️ スライド分析詳細"])
                
                with tab1:
                    st.subheader(f"📊 総合スコア: {result['total_score']} 点")
                    
                    cols = st.columns(4)
                    for i, (label, score) in enumerate(result['sub_scores'].items()):
                        cols[i].caption(f"{label}: {score}点")
                    
                    st.markdown("---")
                    st.markdown(result['evaluation'])
                    
                with tab2:
                    st.text_area("文字起こし内容", result['transcription'], height=300)
                    
                with tab3:
                    st.markdown("### スライドテキスト評価")
                    st.write(result['slide_text_analysis'])
                    if result['image_files']:
                        st.markdown("### 抽出された画像とAIコメント")
                        for img in result['image_files']:
                            if os.path.exists(img):
                                st.image(img, width=300)

            except Exception as e:
                st.error(f"分析中にエラーが発生しました: {e}")
            
            finally:
                # クリーンアップ
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)

        else:
            st.info("音声ファイルとPowerPointファイルを両方アップロードして、分析ボタンを押してください。")

    # フッター
    st.markdown("---")
    st.caption(f"Presentation Evaluator Pro v2.0 (統合版) | Powered by {MODEL_LLM}")


# ==== メイン実行部 ====
if __name__ == "__main__":
    # コマンドライン引数があればCLIモード、なければGUIモード
    if len(sys.argv) > 1:
        # CLIモードで実行
        run_cli_mode()
    else:
        # GUIモードで実行（Streamlitから起動される想定）
        if STREAMLIT_AVAILABLE:
            run_gui_mode()
        else:
            print("エラー: Streamlitがインストールされていません")
            print("GUIモードを使用するには: pip install streamlit")
            print("  streamlit run .\presentation_evaluator.py")
            print("\nCLIモードで使用する場合:")
            print("  python presentation_evaluator.py <音声ファイル> <PowerPointファイル>")
            sys.exit(1)

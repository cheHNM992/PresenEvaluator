# Presentation evaluator
プレゼンテーションを評価するAIプログラム

## 環境構築方法
> git clone https://github.com/cheHNM992/PresenEvaluator.git
> cd .\PresenEvaluator\
> py -3.10 -m venv .venv
> .\.venv\Scripts\activate
> pip install -r requirements.txt

## 実行方法
UI使用の場合
> streamlit run .\presentation_evaluator.py
ブラウザ起動後、OpenAI API Keyを入力し、音声ファイルとパワポファイルを指定、実行

CLI使用の場合
> python .\presentation_evaluator.py <音声ファイル> <PowerPointファイル>


# Presentation evaluator
# プレゼンテーションを評価するAIプログラム
プレゼンテーション用のパワポ資料と、プレゼンを録音した音声ファイルから、プレゼンを総合的に評価するプログラムです。

## 環境構築方法
Pythonのバージョンは3.10.11を推奨します。

### Windowsの場合。
> git clone https://github.com/cheHNM992/PresenEvaluator.git
> cd .\PresenEvaluator\
> py -3.10 -m venv .venv
> .\.venv\Scripts\activate
> pip install -r requirements.txt

### Linuxの場合（恐らくMacも）
> source .venv/bin/activate

## 実行方法
### UI使用の場合
> streamlit run .\presentation_evaluator.py
ブラウザ起動後、OpenAI API Keyを入力し、音声ファイルとパワポファイルを指定、実行。

### CLI使用の場合
set_OpenAI_APIKey.ps1 (Linuxならset_OpenAI_APIKey.sh) にOpenAI API Keyを入力し、下記を実行。
> python .\presentation_evaluator.py <音声ファイル> <PowerPointファイル>

カレントディレクトリに、音声の文字起こしファイルと分析結果ファイルが出力されます。

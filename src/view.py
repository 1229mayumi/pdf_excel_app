from flask import Flask, request, send_file, render_template
# 'pdf2image'ライブラリの'convert_from_bytes'関数をインポート（PDFファイルを画像に変換するために使用）
from pdf2image import convert_from_bytes
# OCRを行うためのライブラリ'pytesseract'をインポート
import pytesseract
# Python Imaging Library（PIL）からImageモジュールをインポート（このライブラリは、画像処理のための機能を提供）
from PIL import Image
import pandas as pd
# 入出力ストリームを扱うためのPython標準ライブラリをインポート（Excelファイルをメモリ内で作成するために使用）
import io
import os

# 'ocr_result'という名前のグローバル変数を初期化
ocr_result = ""
# Flaskのインスタンス化
app = Flask(__name__)

#ルートディレクトリにアクセスがあった場合の挙動
@app.route('/')
def index():
        return render_template("index.html")

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    # 取得したファイルをサーバー上の'uploaded.pdf'というファイル名で保存
    file.save("uploaded.pdf")
    # ファイルのアップロードに成功したとき表示する
    return "ファイルがアップロードされました。"

# OCR1ボタンが押された時の挙動
@app.route('/ocr1', methods=['POST'])
# ocr1という関数を定義（/ocr1へのPOSTリクエストがあった際(ocr1ボタンが押された時）に実行する）
def ocr1():
    global ocr_result
    
    outpdfpass="./tmp.pdf"
    pdf = request.files['file']
    pdf.save(outpdfpass)
    
    # 'uploaded.pdf'という名前のファイルをバイナリ読み取りモード'rb'で開く
    # 'with'ステートメントを使用することで、ファイルの使用が終わった後に自動的に閉じることができる
    with open(outpdfpass, "rb") as file:
        # 開いたファイルから全てのデータを読み取り、それを'pdf_bytes'という変数に保存
        pdf_bytes = file.read()
    
    # PDFを画像に変換して、その結果を'image'という変数に保存
    images = convert_from_bytes(pdf_bytes)

    # 各画像に対してOCRを実行してtext(テキスト)に追加
    text = ''
    # 'text'という変数（変換した画像）に対してOCR処理を行うループを開始
    for image in images:
        text += pytesseract.image_to_string(image)
        
    # OCR処理が終了した後、結果をテキストファイルとして保存
    with open(FILE_PATH, "w", encoding="utf-8") as file:
        file.write(text)
    
    # textに追加された文字列を改行('\n')で分割して、その結果をpandasのDataFrameに変換
    df = pd.DataFrame([text.split('\n')])

    # BytesIO（バイトデータを扱うためのオブジェクト）オブジェクトを作成。直後の行でエクセルファイルを書き込むために使用。
    excel_file = io.BytesIO()
    
    # 書き込み先は'excel_file'、エンジンとして'xlsxwriter：Excelの.xlsxファイルを書き込むためのライブラリ'を指定、indexは書き込まない
    df.to_excel(excel_file, engine='xlsxwriter', index=False)
    excel_file.seek(0)
    
    # 抽出したデータをグローバル変数'ocr_result'に保存
    ocr_result = text 

    os.remove(outpdfpass) 
    
    return "OCR処理は正常に終了しました。"

# 'ocr_result.txt'の絶対パスを取得
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(BASE_DIR, 'ocr_result.txt')
print(BASE_DIR)
print(FILE_PATH)
# 新しいルートの追加
@app.route('/download_ocr_result', methods=['GET'])
def download_ocr_result():
    return send_file(FILE_PATH, as_attachment=True, download_name="ocr_result.txt")

# @app.route('/results', methods=['GET'])
# # ocr結果を表示するresults.htmlへ移動
# def display_results():
#     return render_template('results.html', results=ocr_result)

# @app.route('/download', methods=['GET'])
# # エクセルファイルのダウンロード実行
# def download_file():
#     # ocr結果を改行で分割し、それを１行のDataFrameに変換
#     df = pd.DataFrame([ocr_result.split('\n')])
    
#     excel_file = io.BytesIO()
#     df.to_excel(excel_file, engine='xlsxwriter', index=False)
#     excel_file.seek(0)
#     return send_file(excel_file, attachment_filename='output.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/ocr2', methods=['POST'])
def ocr2():
    # OCR2の機能を実装
    # ...
    return 'OCR2 completed'

@app.route('/ocr3', methods=['POST'])
def ocr3():
    # OCR3の機能を実装
    # ...
    return 'OCR3 completed'

@app.route('/ocr4', methods=['POST'])
def ocr4():
    # OCR4の機能を実装
    # ...
    return 'OCR4 completed'

#エントリーポイント
if __name__ == "__main__":
    app.run(debug=True, port=5000)
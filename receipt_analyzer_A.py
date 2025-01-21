"""
領収証アナライザA

このスクリプトは、りんご領収証PDFファイルを読み込み、抽出したデータをCSVファイルに保存します。
また、CSVを"注文日"・"注文番号"の順で昇順にソートし、"input"フォルダに保存された複数のPDFファイルを順次読み込みます。
さらに、"output"フォルダにリネームしたPDFファイルを保存します。
リネーム形式は以下の通りです：
yyyy-mm-dd_領収証_仕入_<店舗名>_<決済額1>円_<決済2><決済額2>円_<決済3><決済額3>円_<注文番号>.pdf
ERROR発生時には、text化したPDFファイルをerror_YYYYMMDD_hhmm.txtに保存します。

CSVに保存するデータ：
- 注文番号
- 請求番号
- 注文日
- 注文月
- 請求日
- 請求金額
- 請求先
- 配送先
- モール名
- リネーム前ファイル名
- リネーム後ファイル名
- 領収証内情報

関数:
- extract_text_from_pdf(pdf_path): PDFファイルからテキストを抽出し、正規化したテキストを返します。
- initialize_data_structure(): データを保存するための初期構造を返します。
- extract_order_number(text): テキストから注文番号を抽出します。
- extract_billing_number(text): テキストから請求番号を抽出します。
- extract_purchase_date_and_month(text): テキストから注文日と注文月を抽出します。
- copyfile_to_output_folder(pdf_path, new_pdf_file): リネーム後のPDFファイルを元のフォルダ構成でoutputフォルダに保存します。
"""

import os
import unicodedata
import shutil
import pypdf
import pandas as pd
import ctypes


def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = pypdf.PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text() + '\n'  # ページからテキストを抽出して追加
            
    # テキストの正規化
    # 領収証によって異なる文字コードが使用されているため
    # 例：「文」Unicodeの「U+6587」（CJK統合漢字）と「⽂」Unicodeの「U+2F8A」（CJK互換漢字）
    unicode_text = unicodedata.normalize('NFKC', text)
    
    return unicode_text


def initialize_data_structure():
    return {
        '注文番号': [],
        '請求番号': [],
        '請求日': [],
        '注文日': [],
        '注文月': [],
        '請求金額': [],
        '請求先': [],
        '配送先': [],
        'モール名': [],
        'リネーム前ファイル名': [],
        'リネーム後ファイル名': [],
        '領収証内情報': []
    }


def extract_order_number(text):
    # "ご注文番号"を抽出
    if 'ご注文番号:' in text:
        order_number = text.split('ご注文番号:')[1].split('\n')[0]
    else:
        order_number = 'ERROR'

    return order_number


def extract_billing_number(text):
    # "ご請求番号"を抽出
    if 'ご請求番号:' in text:
        billing_number = text.split('ご請求番号:')[1].split('\n')[0]
    else:
        order_number = 'ERROR'

    return billing_number


def extract_purchase_date_and_month(text):
    # 注文日＝"発⾏⽇: mm/dd/yyyy\n"からyyyy/mm/dd形式で抽出
    if '発行日:' in text:
        purchase_data = text.split('発行日:')[1].split('\n')[0]
        purchase_date = purchase_data.split('/')[2] + '/' + purchase_data.split('/')[0] + '/' + purchase_data.split('/')[1]
        purchase_month = purchase_data.split('/')[2] + '/' + purchase_data.split('/')[0]    # 購入日から"yyyy/mm"形式で抽出
    else:
        purchase_date = 'ERROR'
        purchase_month = 'ERROR'

    return purchase_date, purchase_month


def extract_billing_date(text):
    # 請求日＝"請求⽇: mm/dd/yyyy\n"からyyyy/mm/dd形式で抽出
    if '請求日:' in text:
        billing_data = text.split('請求日:')[1].split('\n')[0]
        billing_date = billing_data.split('/')[2] + '/' + billing_data.split('/')[0] + '/' + billing_data.split('/')[1]    # 購入日から"yyyy/mm"形式で抽出
    else:
        billing_date = 'ERROR'

    return billing_date      


def extract_total_price(text):
    # "ご請求⾦額: 219,800 円\n"から"219800"（例）を抽出
    if 'ご請求金額:' in text and '円' in text:
        total_price = text.split('ご請求金額:')[1].split('円')[0]
    else:
        total_price = 'ERROR'

    return total_price


def extract_billing_name(text):
    # "請求先"の次の行の"様"から改行までを請求先として抽出
    if '請求先' in text and '様' in text:
        billing_name = text.split('請求先')[1].split('\n')[5].split('様')[0]
    else:
        billing_name = 'ERROR'

    return billing_name


def extract_shipping_name(text):
    # "配送先"の次の行の"様"から改行までを請求先として抽出
    if '配送先' in text and '様' in text:
        shipping_name = text.split('配送先')[1].split('\n')[5].split('様')[0]
    else:
        shipping_name = 'ERROR'

    return shipping_name

      
def parse_pdf_text(text):
    isError = False
    
    # CSVに保存するデータを初期化
    data = initialize_data_structure()

    # "領収証"から"Page"を抽出 　領収証内情報に絞る作業
    if '領収証' in text and 'Page' in text:
        receipt_text = text.split('領収証')[1].split('Page')[0].replace(' ','')
    else:
        receipt_text = 'ERROR'
        isError = True

    data['領収証内情報'].append(receipt_text)
        
    order_number = extract_order_number(receipt_text)
    data['注文番号'].append(order_number)

    billing_number = extract_billing_number(receipt_text)
    data['請求番号'].append(billing_number)

    purchase_date, purchase_month = extract_purchase_date_and_month(receipt_text)
    data['注文日'].append(purchase_date)
    data['注文月'].append(purchase_month)

    billing_date = extract_billing_date(receipt_text)
    data['請求日'].append(billing_date)

    total_price = extract_total_price(receipt_text)
    data['請求金額'].append(total_price)

    billing_name = extract_billing_name(receipt_text)
    data['請求先'].append(billing_name)
    
    shipping_name = extract_shipping_name(receipt_text)    
    data['配送先'].append(shipping_name)
        
    # モール名をCSVに保存
    data['モール名'].append('AppleJapan合同会社')
    
    isError = False
    if all(value == "ERROR" for value in (order_number, billing_number, purchase_date, billing_date, total_price, billing_name, shipping_name)):
        isError = True
        
    return data, isError


def generate_new_pdf_file_name(data):
    # リネーム後ファイル名の形式：yyyy-mm-dd_領収証_仕入_<店舗名>_<決済額1>円_<決済2><決済額2>円_<決済3><決済額3>円_<注文番号>.pdf
    if 'ERROR' in data['注文番号'][-1] or 'ERROR' in data['請求番号'][-1] or 'ERROR' in data['注文日'][-1] or 'ERROR' in data['請求金額'][-1]:
        return 'ERROR'
    
    new_pdf_file = data['注文日'][-1].replace('/', '-') + '_領収証_仕入_' + data['モール名'][-1] + '_' + data['請求金額'][-1] + '円'

    new_pdf_file += '_注文番号' + data['注文番号'][-1] + '_請求番号' + data['請求番号'][-1] +'.pdf'
    
    return new_pdf_file


def copyfile_to_output_folder(pdf_path, new_pdf_file):
    output_folder = 'output'
    input_folder = 'input'
    relative_path = os.path.relpath(pdf_path, input_folder)
    output_path = os.path.join(output_folder, os.path.dirname(relative_path))

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    if new_pdf_file != 'ERROR': 
        shutil.copyfile(pdf_path, os.path.join(output_path, new_pdf_file))
        

def sort_data(data):
    # データを"購入日"・"注文番号"の順でソート
    df = pd.DataFrame(data)
    df.sort_values(by=['注文日', '注文番号'], inplace=True)
    return df.to_dict(orient='list')


def save_to_csv(data, output_file):
    # データをCSVファイルに保存
    df = pd.DataFrame(data)
    df.to_csv(output_file, index=False, encoding='utf-8_sig')


def main():
    input_folder = 'input'
    pdf_text_with_error = ''
    
    # ユーザーにリネームしたPDFファイルを複製するか尋ねる
    user_response = ctypes.windll.user32.MessageBoxW(0, "リネーム.pdfを複製しますか？\n\n「はい」を選ぶとリスト作成とPDF複製の両方を行います\n「いいえ」を選ぶとリスト作成のみ行います", "確認", 4)
    copy_file_select = 'Yes' if user_response == 6 else 'No'
    
    # 処理中メッセージを表示
    print('処理中... このウィンドウを閉じないでください。')

    # CSVに保存するデータを初期化
    all_data = initialize_data_structure()

    # input_folder直下のPDFファイルを取得
    pdf_files = [file for file in os.listdir(input_folder) if file.endswith('.pdf')]

    # input_folder直下のサブフォルダを取得
    # サブフォルダは1階層のみ対応
    subfolders = [folder for folder in os.listdir(input_folder) 
                  if os.path.isdir(os.path.join(input_folder, folder))]

    # サブフォルダ内のPDFファイルを取得
    for subfolder in subfolders:
        subfolder_path = os.path.join(input_folder, subfolder)
        subfolder_pdf_files = [file for file in os.listdir(subfolder_path) if file.endswith('.pdf')]
        pdf_files.extend([os.path.join(subfolder, file) for file in subfolder_pdf_files])
        
    # 各PDFファイルを順次読み込んでデータを抽出
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file) # PDFファイルのパスを取得
        text = extract_text_from_pdf(pdf_path)  # PDFファイルからテキストを抽出
        data, isError = parse_pdf_text(text) # テキストから必要なデータを抽出
        if isError:
            # ERROR情報を順次追加
            pdf_text_with_error += pdf_file + '\n' + text + '\n\n'
        data['リネーム前ファイル名'].append(pdf_file)   # リネーム前ファイル名をCSVに保存
        new_pdf_file = generate_new_pdf_file_name(data) # リネーム後ファイル名を生成
        data['リネーム後ファイル名'].append(new_pdf_file)   # リネーム後ファイル名をCSVに保存
        
        # outputフォルダにリネームしたPDFファイルをコピーして保存
        if copy_file_select == 'Yes':
            copyfile_to_output_folder(pdf_path, new_pdf_file)
            
        for key in all_data:
            all_data[key].extend(data[key])

    # データをソートする
    sorted_data = sort_data(all_data)

    timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
    # 現在日時を語尾につけてCSVファイルに保存(list_YYYYMMDD_hhmm.csv)
    output_file = 'list_' + timestamp + '.csv'
    save_to_csv(sorted_data, output_file)
    # ERROR発生時にtext化したPDFファイルをerror_YYYYMMDD_hhmm.txtに保存する
    if pdf_text_with_error != '':
        error_file = 'error_' + timestamp + '.txt'
        with open(error_file, 'w', encoding='utf-8_sig') as f:
            f.write(pdf_text_with_error)  
            
    # 処理完了メッセージを表示
    print('...処理が完了しました。')
    # 処理完了をメッセージボックスで通知
    ctypes.windll.user32.MessageBoxW(0, "処理が完了しました。", "完了", 0)

if __name__ == "__main__":
    main()
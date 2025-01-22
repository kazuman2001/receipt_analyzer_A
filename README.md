# 領収証アナライザA

## 概要
領収証アナライザAは、りんご領収証PDFファイルを読み込み、抽出したデータをCSVファイルに保存するスクリプトです。CSVファイルは"注文日"・"注文番号"の順で昇順にソートされます。また、"input"フォルダに保存された複数のPDFファイルを順次読み込み、"output"フォルダにリネームしたPDFファイルを保存します。

## リネーム形式
リネーム後のファイル名は以下の形式で保存されます：

yyyy-mm-dd_領収証_仕入_<店舗名>_<請求額>円_<注文番号>_<請求番号>.pdf


## エラーハンドリング
ERROR発生時には、テキスト化したPDFファイルを`error_YYYYMMDD_hhmm.txt`に保存します。

## CSVに保存するデータ
- 注文番号
- 請求番号
- 注文日
- 注文月
- 請求日
- 請求金額
- 請求先
- 配送先
- モール名
- ギフトカード利用枚数
- ギフトカード利用合計金額
- クレジットカード名
- クレジットカード利用金額
- リネーム前ファイル名
- リネーム後ファイル名
- 領収証内情報

## 使用方法
1. `input`フォルダに領収証PDFファイルを配置します。
2. スクリプトを実行します。
3. 処理が完了すると、`output`フォルダにリネームされたPDFファイルが保存され、抽出されたデータがCSVファイルに保存されます。
4. 'input'フォルダ内にフォルダごと配置することもできます。

## 関数
- `extract_text_from_pdf(pdf_path)`: PDFファイルからテキストを抽出し、正規化したテキストを返します。
- `initialize_data_structure()`: データを保存するための初期構造を返します。
- `extract_order_number(text)`: テキストから注文番号を抽出します。
- `extract_billing_number(text)`: テキストから請求番号を抽出します。
- `extract_purchase_date_and_month(text)`: テキストから注文日と注文月を抽出します。
- `copyfile_to_output_folder(pdf_path, new_pdf_file)`: リネーム後のPDFファイルを元のフォルダ構成で`output`フォルダに保存します。
- `sort_data(data)`: データを"購入日"・"注文番号"の順でソートします。
- `save_to_csv(data, output_file)`: データをCSVファイルに保存します。
- `main()`: スクリプトのメイン処理を行います。

## 注意事項
- サブフォルダは1階層のみ対応しています。
- リネーム後のファイル名に`ERROR`が含まれる場合、そのファイルはリネームされません。

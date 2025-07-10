import os
import csv
import datetime
from card_create import ProductCardGenerator
from page_create import CardPlacementInterface, PageCreator
import tkinter as tk
from tkinter import filedialog
import chardet

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    return result['encoding']

def main():
    # CSVファイルを1回だけ選択
    csv_path = filedialog.askopenfilename(title="CSVファイルを選択", filetypes=[("CSV files", "*.csv")])
    if not csv_path:
        print("CSVファイルが選択されませんでした。")
        return

    # 出力ディレクトリの作成
    output_dir = "output"
    a4_output_dir = "a4_output"
    log_dir = "log"
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(a4_output_dir, exist_ok=True)
    os.makedirs(log_dir, exist_ok=True)

    # エンコーディング自動検出
    encoding = detect_encoding(csv_path)
    if not encoding:
        encoding = 'shift_jis'  # フォールバック

    # CSVデータを読み込み
    with open(csv_path, 'r', encoding=encoding) as f:
        csv_data = list(csv.reader(f))
    header = csv_data[0]
    data_rows = csv_data[1:]

    # タイムスタンプの生成（両方のログで同じものを使用）
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")

    # 1. カード生成プロセス
    generator = ProductCardGenerator()
    card_rows = [header]
    if len(header) < 8 or header[7] != "ステータス":
        header.extend(["ステータス", "エラーログ"])

    # カード生成の処理
    card_results = generator.process_csv_data(data_rows, output_dir)
    card_rows.extend(card_results)

    # カード生成ログの保存
    card_log_filename = f"{timestamp}_card_create_log.csv"
    card_log_path = os.path.join(log_dir, card_log_filename)
    with open(card_log_path, 'w', encoding='shift_jis', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(card_rows)
    print(f"カード作成ログを保存しました: {card_log_path}")

    # 2. ページ生成プロセス
    creator = PageCreator()
    page_rows = [header]

    # ページ生成の処理
    page_results = creator.process_csv_data(data_rows, output_dir, a4_output_dir)
    page_rows.extend(page_results)

    # ページ生成ログの保存
    page_log_filename = f"{timestamp}_page_create_log.csv"
    page_log_path = os.path.join(log_dir, page_log_filename)
    with open(page_log_path, 'w', encoding='shift_jis', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(page_rows)
    print(f"ページ作成ログを保存しました: {page_log_path}")

if __name__ == "__main__":
    main() 
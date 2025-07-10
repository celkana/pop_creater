import os
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
import sys
import datetime
import shutil
from concurrent.futures import ThreadPoolExecutor
import functools
import chardet

class ProductCardGenerator:
    def __init__(self):
        # オールサイズ
        self.full_card_width = 1704
        self.full_card_height = 2000
        self.full_product_image_width = 1582
        self.full_product_image_height = 1582
        self.full_product_image_x = 57
        self.full_product_image_y = 60
        self.full_number_x = 852  # センター（カラー英字と同じ）
        self.full_number_y = 1956
        self.full_product_name_x = 852  # センター
        self.full_product_name_y = 1710  # 画像下部（95+1582=1677）から少し下
        self.full_color_jp_x = 852  # センター
        self.full_color_jp_y = 1829
        self.full_color_en_x = 852  # センター
        self.full_color_en_y = 1893
        self.full_font_normal_size = 50  # 12pt → 50px
        self.full_font_large_size = 83   # 20pt → 83px
        self.full_font_number_size = 33  # 8pt → 33px
        
        # ハーフサイズ
        self.half_card_width = 852
        self.half_card_height = 1000
        self.half_product_image_width = 730
        self.half_product_image_height = 730
        self.half_product_image_x = 60
        self.half_product_image_y = 60
        self.half_number_x = 426  # センター（カラー英字と同じ）
        self.half_number_y = 956  # カラー英字（y=896）の下（896+33+10）
        self.half_product_name_x = 426  # センター
        self.half_product_name_y = 790
        self.half_color_jp_x = 426  # センター
        self.half_color_jp_y = 849
        self.half_color_en_x = 426  # センター
        self.half_color_en_y = 896 
        self.half_font_normal_size = 42  # 10pt → 42px
        self.half_font_large_size = 67   # 16pt → 67px
        self.half_font_number_size = 33  # 8pt → 33px

        # フォントの初期化（キャッシュ）
        self.fonts = self._initialize_fonts()
        
        # 画像キャッシュ
        self.image_cache = {}

    def _initialize_fonts(self):
        fonts = {}
        try:
            font_path = "C:/Windows/Fonts/YUGOTHR.TTC" if sys.platform == "win32" else "/usr/share/fonts/truetype/yu/YuGothic-Regular.ttf"
            font_path_bold = "C:/Windows/Fonts/YUGOTHB.TTC" if sys.platform == "win32" else "/usr/share/fonts/truetype/yu/YuGothic-Bold.ttf"
            
            # フルサイズ用フォント
            fonts['full_normal'] = ImageFont.truetype(font_path, self.full_font_normal_size, index=0)
            fonts['full_large'] = ImageFont.truetype(font_path_bold, self.full_font_large_size, index=1)
            fonts['full_number'] = ImageFont.truetype(font_path, self.full_font_number_size, index=0)
            
            # ハーフサイズ用フォント
            fonts['half_normal'] = ImageFont.truetype(font_path, self.half_font_normal_size, index=0)
            fonts['half_large'] = ImageFont.truetype(font_path_bold, self.half_font_large_size, index=1)
            fonts['half_number'] = ImageFont.truetype(font_path, self.half_font_number_size, index=0)
        except:
            try:
                # フォールバックフォント
                fonts['full_normal'] = ImageFont.truetype("msgothic.ttc", self.full_font_normal_size)
                fonts['full_large'] = ImageFont.truetype("msgothic.ttc", self.full_font_large_size)
                fonts['full_number'] = ImageFont.truetype("msgothic.ttc", self.full_font_number_size)
                fonts['half_normal'] = ImageFont.truetype("msgothic.ttc", self.half_font_normal_size)
                fonts['half_large'] = ImageFont.truetype("msgothic.ttc", self.half_font_large_size)
                fonts['half_number'] = ImageFont.truetype("msgothic.ttc", self.half_font_number_size)
            except:
                # 最終フォールバック
                default_font = ImageFont.load_default()
                fonts = {k: default_font for k in ['full_normal', 'full_large', 'full_number', 
                                                 'half_normal', 'half_large', 'half_number']}
        return fonts

    def _get_cached_image(self, image_path):
        if image_path not in self.image_cache:
            try:
                image = Image.open(image_path)
                self.image_cache[image_path] = image
            except Exception as e:
                print(f"画像読み込みエラー: {e}")
                return None
        return self.image_cache[image_path]

    def fit_to_rect(self, image, width, height):
        # 画像のアスペクト比を計算
        img_ratio = image.width / image.height
        target_ratio = width / height

        if img_ratio > target_ratio:
            # 画像が横長の場合、幅に合わせる
            new_width = width
            new_height = int(width / img_ratio)
        else:
            # 画像が縦長の場合、高さに合わせる
            new_height = height
            new_width = int(height * img_ratio)

        # 画像をリサイズ
        resized = image.resize((new_width, new_height), Image.Resampling.LANCZOS)

        # 背景を作成
        background = Image.new(image.mode, (width, height), 'white' if image.mode == 'RGB' else (255, 255, 255, 0))

        # 中央に配置するためのオフセットを計算
        offset_x = (width - new_width) // 2
        offset_y = (height - new_height) // 2

        # 画像を中央に配置
        background.paste(resized, (offset_x, offset_y))
        return background

    def create_card(self, image_path, product_name, color_jp, color_en, item_code, size="full"):
        # サイズに応じたパラメータの設定
        params = {
            "full": {
                "card_width": self.full_card_width,
                "card_height": self.full_card_height,
                "product_image_width": self.full_product_image_width,
                "product_image_height": self.full_product_image_height,
                "product_image_x": self.full_product_image_x,
                "product_image_y": self.full_product_image_y,
                "number_x": self.full_number_x,
                "number_y": self.full_number_y,
                "product_name_x": self.full_product_name_x,
                "product_name_y": self.full_product_name_y,
                "color_jp_x": self.full_color_jp_x,
                "color_jp_y": self.full_color_jp_y,
                "color_en_x": self.full_color_en_x,
                "color_en_y": self.full_color_en_y,
                "fonts": {
                    "normal": self.fonts['full_normal'],
                    "large": self.fonts['full_large'],
                    "number": self.fonts['full_number']
                }
            },
            "half": {
                "card_width": self.half_card_width,
                "card_height": self.half_card_height,
                "product_image_width": self.half_product_image_width,
                "product_image_height": self.half_product_image_height,
                "product_image_x": self.half_product_image_x,
                "product_image_y": self.half_product_image_y,
                "number_x": self.half_number_x,
                "number_y": self.half_number_y,
                "product_name_x": self.half_product_name_x,
                "product_name_y": self.half_product_name_y,
                "color_jp_x": self.half_color_jp_x,
                "color_jp_y": self.half_color_jp_y,
                "color_en_x": self.half_color_en_x,
                "color_en_y": self.half_color_en_y,
                "fonts": {
                    "normal": self.fonts['half_normal'],
                    "large": self.fonts['half_large'],
                    "number": self.fonts['half_number']
                }
            }
        }[size]

        # カード作成
        card = Image.new('RGB', (params["card_width"], params["card_height"]), 'white')
        draw = ImageDraw.Draw(card)

        # 商品画像
        product_image = self._get_cached_image(image_path)
        if product_image is None:
            return None

        if product_image.mode == 'RGBA':
            product_image = self.fit_to_rect(
                product_image.convert('RGBA'),
                params["product_image_width"],
                params["product_image_height"]
            )
            card = card.convert('RGBA')
            card.paste(product_image, (params["product_image_x"], params["product_image_y"]), product_image.split()[3])
            card = card.convert('RGB')
        else:
            product_image = self.fit_to_rect(
                product_image.convert('RGB'),
                params["product_image_width"],
                params["product_image_height"]
            )
            card.paste(product_image, (params["product_image_x"], params["product_image_y"]))

        # テキスト描画
        draw.text(
            (params["number_x"], params["number_y"]),
            item_code,
            font=params["fonts"]["number"],
            fill='black',
            stroke_width=5,
            stroke_fill='white',
            anchor="mm"
        )
        draw.text(
            (params["product_name_x"], params["product_name_y"]),
            product_name,
            font=params["fonts"]["normal"],
            fill='black',
            stroke_width=5,
            stroke_fill='white',
            anchor="mm"
        )
        draw.text(
            (params["color_en_x"], params["color_en_y"]),
            color_en,
            font=params["fonts"]["large"],
            fill='black',
            stroke_width=5,
            stroke_fill='white',
            anchor="mm"
        )
        draw.text(
            (params["color_jp_x"], params["color_jp_y"]),
            color_jp,
            font=params["fonts"]["normal"],
            fill='black',
            stroke_width=5,
            stroke_fill='white',
            anchor="mm"
        )

        return card

    def process_csv(self):
        csv_path = filedialog.askopenfilename(title="CSVファイルを選択", filetypes=[("CSV files", "*.csv")])
        if not csv_path:
            print("CSVファイルが選択されませんでした。")
            return None

        # エンコーディング自動検出
        encoding = detect_encoding(csv_path)
        if not encoding:
            encoding = 'shift_jis'  # フォールバック

        # CSVを読み込み
        rows = []
        with open(csv_path, 'r', encoding=encoding) as f:
            reader = csv.reader(f)
            header = next(reader)
            if len(header) < 8 or header[7] != "ステータス":
                header.extend(["ステータス", "エラーログ"])
            rows.append(header)

            data_rows = list(reader)

        # 画像生成の並列処理
        with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
            process_func = functools.partial(process_row, self, output_dir="output")
            results = list(executor.map(lambda row: process_func(row), data_rows))

        rows.extend(results)
        return rows

    def process_csv_data(self, data_rows, output_dir):
        results = []
        for row in data_rows:
            try:
                if len(row) < 5:  # 必要な列数をチェック
                    results.append(row + ["エラー", "データ列が不足しています"])
                    continue

                item_code = row[0]
                image_path = row[1]
                product_name = row[2]
                color_jp = row[3]
                color_en = row[4]

                status = "正常"
                error_log = ""

                if not os.path.exists(image_path):
                    results.append(row + ["エラー", f"画像ファイルが見つかりません: {image_path}"])
                    continue

                # フルサイズカード生成
                full_card = self.create_card(
                    image_path, product_name, color_jp, color_en, item_code, size="full"
                )
                if full_card:
                    output_path = os.path.join(output_dir, f"{item_code}-{product_name}_{color_jp}_full.png")
                    full_card.save(output_path)
                    print(f"カード（オール）を生成しました: {output_path}")
                else:
                    status = "エラー"
                    error_log = "フルサイズカードの生成に失敗しました"

                # ハーフサイズカード生成
                half_card = self.create_card(
                    image_path, product_name, color_jp, color_en, item_code, size="half"
                )
                if half_card:
                    output_path = os.path.join(output_dir, f"{item_code}-{product_name}_{color_jp}_half.png")
                    half_card.save(output_path)
                    print(f"カード（ハーフ）を生成しました: {output_path}")
                else:
                    status = "エラー"
                    error_log += ", ハーフサイズカードの生成に失敗しました" if error_log else "ハーフサイズカードの生成に失敗しました"

                results.append(row + [status, error_log])
            except Exception as e:
                results.append(row + ["エラー", f"処理中にエラーが発生: {str(e)}"])

        return results

    def save_log(self, results, log_dir="log"):
        # logフォルダが存在しない場合は作成
        os.makedirs(log_dir, exist_ok=True)
        
        # 現在の日時を取得してファイル名を作成
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        log_filename = f"{timestamp}_card_create_log.csv"
        log_path = os.path.join(log_dir, log_filename)
        
        # CSVに保存（shift-jisエンコーディングで保存）
        with open(log_path, 'w', encoding='shift_jis', newline='') as f:
            writer = csv.writer(f)
            # データを書き込む
            writer.writerows(results)
        
        print(f"カード作成ログを保存しました: {log_path}")
        return log_path

def process_row(generator, row, output_dir):
    try:
        item_code = row[0]
        image_path = row[1]
        product_name = row[2]
        color_jp = row[3]
        color_en = row[4]
        page_product_name = row[5]
        
        status = "正常"
        error_log = ""
        
        if not os.path.exists(image_path):
            return row + ["エラー", f"画像ファイルが見つかりません: {image_path}"]
            
        # フルサイズカード生成
        full_card = generator.create_card(
            image_path, product_name, color_jp, color_en, item_code, size="full"
        )
        if full_card:
            output_path = os.path.join(output_dir, f"{item_code}-{product_name}_{color_jp}_full.png")
            full_card.save(output_path)
            print(f"カード（オール）を生成しました: {output_path}")
        else:
            status = "エラー"
            error_log = "フルサイズカードの生成に失敗しました"
            
        # ハーフサイズカード生成
        half_card = generator.create_card(
            image_path, product_name, color_jp, color_en, item_code, size="half"
        )
        if half_card:
            output_path = os.path.join(output_dir, f"{item_code}-{product_name}_{color_jp}_half.png")
            half_card.save(output_path)
            print(f"カード（ハーフ）を生成しました: {output_path}")
        else:
            status = "エラー"
            error_log += ", ハーフサイズカードの生成に失敗しました" if error_log else "ハーフサイズカードの生成に失敗しました"
            
        return row + [status, error_log]
    except Exception as e:
        return row + ["エラー", f"処理中にエラーが発生: {str(e)}"]

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    return result['encoding']

def main():
    csv_path = filedialog.askopenfilename(title="CSVファイルを選択", filetypes=[("CSV files", "*.csv")])
    if not csv_path:
        print("CSVファイルが選択されませんでした。")
        return
        
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    
    log_dir = "log"
    os.makedirs(log_dir, exist_ok=True)
    
    # エンコーディング自動検出
    encoding = detect_encoding(csv_path)
    if not encoding:
        encoding = 'shift_jis'  # フォールバック
    
    # CSVを読み込み
    rows = []
    with open(csv_path, 'r', encoding=encoding) as f:
        reader = csv.reader(f)
        header = next(reader)
        # ヘッダーを修正
        header = ["商品コード", "画像パス", "商品名", "カラー名", "カラー名(英語)", "ページ商品名", "ステータス", "エラーログ"]
        rows.append(header)
        
        data_rows = list(reader)
    
    # 画像生成の並列処理
    generator = ProductCardGenerator()
    with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
        process_func = functools.partial(process_row, generator, output_dir=output_dir)
        results = list(executor.map(lambda row: process_func(row), data_rows))
    
    rows.extend(results)
    
    # ログファイル保存
    log_path = generator.save_log(rows, log_dir=log_dir)  # resultsではなくrowsを渡す
    
    # キャッシュクリア
    generator.image_cache.clear()

if __name__ == "__main__":
    main()
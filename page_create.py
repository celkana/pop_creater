import os
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont, ImageTk
import sys
import datetime
from collections import defaultdict
import chardet
from card_layouts import CardLayoutManager

class A4PageGenerator:
    def __init__(self):
        # A4サイズ（300dpi、横）
        self.a4_width = 3508
        self.a4_height = 2480
        
        # ヘッダーサイズ
        self.header_width = 3408
        self.header_height = 270
        
        # ヘッダーの位置（中央寄せ）
        self.header_x = int((self.a4_width - self.header_width) / 2)  # 50px
        self.header_y = 60
        self.header_border_width = 10
        
        # フォントサイズ（20pt = 86px @ 300dpi）
        self.header_font_size = 86

    def create_a4_page(self, cards, page_product_name):
        page = Image.new('RGB', (self.a4_width, self.a4_height), 'white')
        draw = ImageDraw.Draw(page)

        # ヘッダー
        draw.rectangle(
            (self.header_x, self.header_y, 
             self.header_x + self.header_width, self.header_y + self.header_height),
            outline='black', width=self.header_border_width
        )

        # フォント
        try:
            font_path = "C:/Windows/Fonts/YUGOTHB.TTC" if sys.platform == "win32" else "/usr/share/fonts/truetype/yu/YuGothic-Bold.ttf"
            font = ImageFont.truetype(font_path, self.header_font_size, index=0)
        except:
            try:
                font = ImageFont.truetype("msgothic.ttc", self.header_font_size)
            except:
                font = ImageFont.load_default()

        # ヘッダーテキスト
        # 他バリエーションはこちら（左揃え）
        draw.text((self.header_x + 50, self.header_y + self.header_height/2), 
                  "他バリエーションはこちら", font=font, fill='black', anchor="lm")
        
        # 商品名（右揃え）
        draw.text((self.header_x + self.header_width - 50, self.header_y + self.header_height/2), 
                  f"商品名：{page_product_name}", font=font, fill='black', anchor="rm")

        # カード配置
        for card_info in cards:
            card_img = card_info['image']
            x = card_info['x']
            y = card_info['y']
            page.paste(card_img, (x, y))

        return page

    def calculate_positions_and_sizes(self, num_cards):
        positions = []
        sizes = []  # 'full' または 'half'
        
        # カードサイズ定義（300dpi）
        full_width = 1704  # フルサイズの幅
        full_height = 2000  # フルサイズの高さ
        half_width = 852   # ハーフサイズの幅
        half_height = 1000 # ハーフサイズの高さ
        
        # 余白
        margin_top = 400
        
        if num_cards == 1:
            positions = [(int((self.a4_width - full_width) / 2), margin_top)]  # 中央
            sizes = ['full']
        elif num_cards == 2:
            # 2枚の場合はフルサイズ2枚を横に並べて中央揃え
            total_width = full_width * 2
            center_offset = int((self.a4_width - total_width) / 2)
            positions = [
                (center_offset, margin_top),
                (center_offset + full_width, margin_top)
            ]
            sizes = ['full', 'full']
        elif num_cards == 3:
            # 全体の幅を計算（フル幅 + ハーフ幅）
            total_width = full_width + half_width
            # 中央寄せのためのオフセットを計算
            center_offset = int((self.a4_width - total_width) / 2)
            
            # フル画像とハーフ画像の位置を計算
            x1 = center_offset  # フル画像の開始位置
            x2 = x1 + full_width  # ハーフ画像の開始位置（フル画像の直後）
            
            positions = [
                (x1, margin_top),  # 左（フル）
                (x2, margin_top),  # 右上（ハーフ）
                (x2, margin_top + half_height)  # 右下（ハーフ）、上のハーフ画像にくっつける
            ]
            sizes = ['full', 'half', 'half']
        elif num_cards == 4:
            # フル画像1枚 + 残りをハーフサイズで配置
            num_half = num_cards - 1
            num_half_cols = 2  # ハーフサイズ画像の列数
            num_half_rows = (num_half + num_half_cols - 1) // num_half_cols  # 切り上げ除算

            # 全体の幅を計算（フル幅 + ハーフ幅 * 列数）
            total_width = full_width + half_width * num_half_cols
            # 中央寄せのためのオフセットを計算
            center_offset = int((self.a4_width - total_width) / 2)

            # フル画像の位置（左側）
            positions.append((center_offset, margin_top))
            sizes.append('full')

            # ハーフ画像の位置（右側に格子状に配置）
            half_start_x = center_offset + full_width
            for i in range(num_half):
                col = i % num_half_cols
                row = i // num_half_cols
                x = half_start_x + col * half_width
                y = margin_top + row * half_height
                positions.append((x, y))
                sizes.append('half')
        elif num_cards == 5:
            # フル画像1枚 + 残りをハーフサイズで配置
            num_half = num_cards - 1
            num_half_cols = 2  # ハーフサイズ画像の列数
            num_half_rows = (num_half + num_half_cols - 1) // num_half_cols  # 切り上げ除算

            # 全体の幅を計算（フル幅 + ハーフ幅 * 列数）
            total_width = full_width + half_width * num_half_cols
            # 中央寄せのためのオフセットを計算
            center_offset = int((self.a4_width - total_width) / 2)

            # フル画像の位置（左側）
            positions.append((center_offset, margin_top))
            sizes.append('full')

            # ハーフ画像の位置（右側に格子状に配置）
            half_start_x = center_offset + full_width
            for i in range(num_half):
                col = i % num_half_cols
                row = i // num_half_cols
                x = half_start_x + col * half_width
                y = margin_top + row * half_height
                positions.append((x, y))
                sizes.append('half')
        elif num_cards == 6:
            # ハーフサイズを3x2で配置
            num_cols = 3
            num_rows = 2
            total_width = half_width * num_cols
            center_offset_x = int((self.a4_width - total_width) / 2)

            for i in range(num_cards):
                col = i % num_cols
                row = i // num_cols
                x = center_offset_x + col * half_width
                y = margin_top + row * half_height
                positions.append((x, y))
                sizes.append('half')
        elif num_cards == 7:
            # ハーフサイズを4+3で配置
            total_width = half_width * 4
            center_offset_x = int((self.a4_width - total_width) / 2)

            # 上段4枚
            for i in range(4):
                x = center_offset_x + i * half_width
                y = margin_top
                positions.append((x, y))
                sizes.append('half')

            # 下段3枚
            total_width = half_width * 3
            center_offset_x = int((self.a4_width - total_width) / 2)
            for i in range(3):
                x = center_offset_x + i * half_width
                y = margin_top + half_height
                positions.append((x, y))
                sizes.append('half')
        elif num_cards == 8:
            # ハーフサイズを4x2で配置
            num_cols = 4
            num_rows = 2
            total_width = half_width * num_cols
            center_offset_x = int((self.a4_width - total_width) / 2)

            for i in range(num_cards):
                col = i % num_cols
                row = i // num_cols
                x = center_offset_x + col * half_width
                y = margin_top + row * half_height
                positions.append((x, y))
                sizes.append('half')

        return positions, sizes

class CardPlacementInterface:
    def __init__(self, page_data=None):
        self.root = tk.Tk()
        self.root.title("カード配置")
        self.root.geometry("1200x800")
        
        # カードデータ
        self.card_data = []
        self.current_page = None
        self.current_page_index = -1
        
        # カード配置
        self.card_positions = []
        self.card_sizes = []
        self.layout_manager = CardLayoutManager()
        self.current_layout = None
        
        # ページジェネレーター
        self.page_generator = A4PageGenerator()
        
        # 外部から渡されたページデータ
        self.page_data = page_data
        self.all_cards = []
        self.page_product_names = []
        
        # ログ用データ
        self.log_data = []
        
        self.setup_gui()
        
    def setup_gui(self):
        # 左側：カードリスト
        left_frame = ttk.Frame(self.root)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        
        # パターン選択フレーム
        pattern_frame = ttk.Frame(left_frame)
        pattern_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        
        # パターン選択ラベル
        ttk.Label(pattern_frame, text="配置パターン:").pack(side=tk.TOP, anchor=tk.W)
        
        # パターン選択用の変数とラジオボタン
        self.pattern_var = tk.StringVar(value="pattern1")
        self.pattern_radios = []  # ラジオボタンを保持
        
        # パターンラジオボタンフレーム（後でupdate_pattern_choicesで更新）
        self.pattern_radio_frame = ttk.Frame(pattern_frame)
        self.pattern_radio_frame.pack(side=tk.TOP, fill=tk.X)
        
        # カードリスト
        self.card_listbox = tk.Listbox(left_frame, width=40, height=30)
        self.card_listbox.pack(side=tk.TOP, fill=tk.Y)
        
        # ボタンフレーム
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        
        # 上下ボタン
        ttk.Button(button_frame, text="↑", command=self.move_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="↓", command=self.move_down).pack(side=tk.LEFT, padx=2)
        
        # 反映・決定・スキップボタン
        ttk.Button(left_frame, text="反映", command=self.reflect_changes).pack(side=tk.TOP, fill=tk.X, pady=5)
        ttk.Button(left_frame, text="決定", command=self.finish_page).pack(side=tk.TOP, fill=tk.X, pady=5)
        ttk.Button(left_frame, text="スキップ", command=self.skip_page).pack(side=tk.TOP, fill=tk.X, pady=5)
        
        # 右側：プレビュー
        right_frame = ttk.Frame(self.root)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # プレビューキャンバス（ウィンドウサイズに合わせてスケール調整）
        self.root.update_idletasks()
        available_width = right_frame.winfo_width()
        available_height = right_frame.winfo_height()
        if available_width == 1 or available_height == 1:  # 初期値が1の場合
            available_width = 600
            available_height = 600
        
        # スケールを動的に計算（ウィンドウにフィットするように）
        scale_x = available_width / self.page_generator.a4_width
        scale_y = available_height / self.page_generator.a4_height
        self.scale = min(scale_x, scale_y) * 0.9  # 少し余裕を持たせる
        
        canvas_width = int(self.page_generator.a4_width * self.scale)
        canvas_height = int(self.page_generator.a4_height * self.scale)
        
        self.canvas = tk.Canvas(
            right_frame, 
            width=canvas_width, 
            height=canvas_height, 
            bg='white'
        )
        self.canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
    def update_pattern_choices(self):
        """利用可能なパターンの選択肢を更新"""
        # 既存のラジオボタンをクリア
        for radio in self.pattern_radios:
            radio.destroy()
        self.pattern_radios.clear()
        
        # 現在のカード枚数に対応するパターンを取得
        layouts = self.layout_manager.get_layouts_for_count(len(self.card_data))
        if not layouts:
            return
        
        # 最初のパターンをデフォルトで選択
        self.pattern_var.set(next(iter(layouts.keys())))
        
        # 各パターンのラジオボタンを作成
        for pattern_id, layout in layouts.items():
            radio = ttk.Radiobutton(
                self.pattern_radio_frame,
                text=f"{layout.name}: {layout.description}",
                value=pattern_id,
                variable=self.pattern_var,
                command=self.reflect_changes
            )
            radio.pack(side=tk.TOP, anchor=tk.W)
            self.pattern_radios.append(radio)

    def move_up(self):
        selected = self.card_listbox.curselection()
        if not selected or selected[0] == 0:
            return
        
        idx = selected[0]
        # リストボックスの項目を入れ替え
        text = self.card_listbox.get(idx)
        self.card_listbox.delete(idx)
        self.card_listbox.insert(idx-1, text)
        self.card_listbox.selection_set(idx-1)
        
        # カードデータも入れ替え
        self.card_data[idx], self.card_data[idx-1] = self.card_data[idx-1], self.card_data[idx]
        
        # プレビューを更新
        self.reflect_changes()
        
    def move_down(self):
        selected = self.card_listbox.curselection()
        if not selected or selected[0] == self.card_listbox.size() - 1:
            return
        
        idx = selected[0]
        # リストボックスの項目を入れ替え
        text = self.card_listbox.get(idx)
        self.card_listbox.delete(idx)
        self.card_listbox.insert(idx+1, text)
        self.card_listbox.selection_set(idx+1)
        
        # カードデータも入れ替え
        self.card_data[idx], self.card_data[idx+1] = self.card_data[idx+1], self.card_data[idx]
        
        # プレビューを更新
        self.reflect_changes()
        
    def load_page(self, page_product_name):
        # カードデータをクリア
        self.card_data = []
        self.card_positions = []
        self.card_sizes = []
        self.current_page = page_product_name
        
        # カードリストをクリア
        self.card_listbox.delete(0, tk.END)
        
        # カードデータを収集
        for card in self.all_cards:
            if card['page_product_name'] == page_product_name:
                self.card_data.append(card)
                self.card_listbox.insert(tk.END, f"{card['product_name']} - {card['color_jp']}")
        
        # パターン選択肢を更新
        self.update_pattern_choices()
        
        # プレビューを更新
        self.reflect_changes()
        
    def reflect_changes(self):
        # キャンバスをクリア
        self.canvas.delete("all")
        self.preview_images = []  # 参照をクリア
        
        # A4用紙の背景を描画
        self.canvas.create_rectangle(
            0, 0, 
            int(self.page_generator.a4_width * self.scale), 
            int(self.page_generator.a4_height * self.scale), 
            fill="white", outline="black"
        )
        
        # ヘッダーを描画
        header_x = int(self.page_generator.header_x * self.scale)
        header_y = int(self.page_generator.header_y * self.scale)
        header_width = int(self.page_generator.header_width * self.scale)
        header_height = int(self.page_generator.header_height * self.scale)
        
        self.canvas.create_rectangle(
            header_x, header_y,
            header_x + header_width, header_y + header_height,
            outline="black", width=int(self.page_generator.header_border_width * self.scale)
        )
        
        # ヘッダーテキスト（プレビュー用の簡易表示）
        self.canvas.create_text(
            header_x + int(50 * self.scale), 
            header_y + int(header_height/2), 
            text="他バリエーションはこちら", 
            anchor="w"
        )
        
        self.canvas.create_text(
            header_x + header_width - int(50 * self.scale), 
            header_y + int(header_height/2), 
            text=f"商品名：{self.current_page if self.current_page else '未選択'}", 
            anchor="e"
        )
        
        # 選択されているパターンのレイアウトを取得
        layouts = self.layout_manager.get_layouts_for_count(len(self.card_data))
        if not layouts:
            return
            
        selected_pattern = self.pattern_var.get()
        layout = layouts.get(selected_pattern)
        if not layout:
            return
            
        self.card_positions = layout.positions
        self.card_sizes = layout.sizes
        
        # カードを描画
        missing_images = []
        
        for i, (pos, size) in enumerate(zip(self.card_positions, self.card_sizes)):
            if i < len(self.card_data):
                card = self.card_data[i]
                image_path = os.path.join(
                    "output",
                    f"{card['item_code']}-{card['product_name']}_{card['color_jp']}_{'full' if size == 'full' else 'half'}.png"
                )
                
                print(f"プレビュー用画像パス: {image_path}")  # デバッグ用
                
                if os.path.exists(image_path):
                    try:
                        image = Image.open(image_path)
                        # 表示用に縮小
                        image = image.resize(
                            (int(image.width * self.scale), int(image.height * self.scale)),
                            Image.Resampling.LANCZOS
                        )
                        photo = ImageTk.PhotoImage(image)
                        # スケールした位置に描画
                        self.canvas.create_image(
                            int(pos[0] * self.scale), 
                            int(pos[1] * self.scale), 
                            image=photo, 
                            anchor=tk.NW
                        )
                        self.preview_images.append(photo)  # 参照をリストで保持
                    except Exception as e:
                        print(f"画像読み込みエラー: {image_path}, エラー: {str(e)}")
                        missing_images.append(f"{card['product_name']} - {card['color_jp']}")
                        self.draw_missing_image_placeholder(pos, size)
                else:
                    print(f"画像が見つかりません: {image_path}")
                    missing_images.append(f"{card['product_name']} - {card['color_jp']}")
                    self.draw_missing_image_placeholder(pos, size)
        
        # 不足している画像があれば警告表示
        if missing_images:
            warning_text = "画像が見つかりません: " + ", ".join(missing_images[:3])
            if len(missing_images) > 3:
                warning_text += f" 他{len(missing_images) - 3}件"
            
            self.canvas.create_rectangle(
                10, 10, 
                int(self.page_generator.a4_width * self.scale) - 10, 
                50, 
                fill="yellow", outline="red"
            )
            self.canvas.create_text(
                20, 30,
                text=warning_text,
                anchor="w",
                fill="red"
            )
    
    def draw_missing_image_placeholder(self, pos, size):
        # 不足している画像の代わりに表示するプレースホルダーを描画
        full_width = int(1704 * self.scale) if size == 'full' else int(852 * self.scale)
        full_height = int(2000 * self.scale) if size == 'full' else int(1000 * self.scale)
        
        x = int(pos[0] * self.scale)
        y = int(pos[1] * self.scale)
        
        # 赤い枠の四角形を描画
        self.canvas.create_rectangle(
            x, y, 
            x + full_width, y + full_height,
            outline="red", width=2
        )
        
        # クロスマークを描画
        self.canvas.create_line(
            x, y, 
            x + full_width, y + full_height,
            fill="red", width=2
        )
        self.canvas.create_line(
            x + full_width, y, 
            x, y + full_height,
            fill="red", width=2
        )
        
        # 「画像なし」テキストを描画
        self.canvas.create_text(
            x + full_width/2, 
            y + full_height/2,
            text="画像なし",
            fill="red",
            font=("Helvetica", 16)
        )
        
    def finish_page(self):
        if not self.current_page:
            return
            
        # ページを保存
        output_dir = "a4_output"
        os.makedirs(output_dir, exist_ok=True)
        
        # カード情報を準備
        cards = []
        missing_images = []
        
        for i, (pos, size) in enumerate(zip(self.card_positions, self.card_sizes)):
            if i < len(self.card_data):
                card = self.card_data[i]
                image_path = os.path.join(
                    "output",
                    f"{card['item_code']}-{card['product_name']}_{card['color_jp']}_{'full' if size == 'full' else 'half'}.png"
                )
                
                # 画像パスを出力（デバッグ用）
                print(f"最終画像用パス: {image_path}")
                
                if os.path.exists(image_path):
                    try:
                        card_image = Image.open(image_path)
                        cards.append({
                            'image': card_image,
                            'x': pos[0],
                            'y': pos[1],
                            'item_code': card['item_code'],
                            'size': size
                        })
                    except Exception as e:
                        print(f"最終画像読み込みエラー: {image_path}, エラー: {str(e)}")
                        missing_images.append(f"{card['product_name']}-{card['color_jp']}")
                else:
                    print(f"最終画像が見つかりません: {image_path}")
                    missing_images.append(f"{card['product_name']}-{card['color_jp']}")
        
        # ページを生成
        status = "正常"
        log_message = ""
        
        if missing_images:
            status = "異常"
            log_message = f"足りない画像があります: {', '.join(missing_images[:3])}"
            if len(missing_images) > 3:
                log_message += f" 他{len(missing_images) - 3}件"
        
        # ログデータを追加
        self.log_data.append({
            "page_name": self.current_page,
            "status": status,
            "log": log_message,
            "card_count": str(len(cards))  # 単純に枚数のみを表示
        })
        
        if cards:
            page_image = self.page_generator.create_a4_page(cards, self.current_page)
            
            # ページを保存（ファイル名に使えない文字を置換）
            safe_page_name = self.current_page.replace('/', '／').replace('\\', '＼').replace(':', '：').replace('*', '＊').replace('?', '？').replace('"', "'").replace('<', '＜').replace('>', '＞').replace('|', '｜')
            output_path = os.path.join(output_dir, f"{safe_page_name}.png")
            page_image.save(output_path)
            print(f"ページを保存しました: {output_path}")
        else:
            print(f"カードがないため、ページ {self.current_page} は生成されませんでした。")
        
        # 次のページへ
        self.go_to_next_page()
    
    def skip_page(self):
        if not self.current_page:
            return
            
        # ログデータを追加
        self.log_data.append({
            "page_name": self.current_page,
            "status": "スキップ",
            "log": "ユーザーによりスキップされました",
            "card_count": "0"  # スキップ時は0を表示
        })
        
        print(f"ページ {self.current_page} をスキップしました。")
        
        # 次のページへ
        self.go_to_next_page()
    
    def go_to_next_page(self):
        # 次のページへ
        self.current_page_index += 1
        if self.current_page_index < len(self.page_product_names):
            self.current_page = self.page_product_names[self.current_page_index]
            self.load_page(self.current_page)
        else:
            # すべてのページが完了したらログを保存
            self.save_log()
            self.root.destroy()
    
    def save_log(self):
        # logフォルダが存在しない場合は作成
        log_dir = "log"
        os.makedirs(log_dir, exist_ok=True)
        
        # 現在の日時を取得してファイル名を作成
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        log_filename = f"{timestamp}_page_create_log.csv"
        log_path = os.path.join(log_dir, log_filename)
        
        # CSVに保存（shift-jisエンコーディングで保存）
        with open(log_path, 'w', encoding='shift_jis', newline='') as f:
            writer = csv.writer(f)
            # ヘッダーを書き込む
            writer.writerow(["ページ名", "ステータス", "ログ", "カード枚数"])
            # データを書き込む
            for log_entry in self.log_data:
                writer.writerow([
                    log_entry["page_name"],
                    log_entry["status"],
                    log_entry["log"],
                    log_entry["card_count"]
                ])
        
        print(f"ログを保存しました: {log_path}")
    
    def main(self):
        # 外部からデータが渡された場合はそれを使用
        if self.page_data:
            self.process_page_data()
        else:
            # CSVファイルを選択
            csv_path = filedialog.askopenfilename(title="CSVファイルを選択", filetypes=[("CSV files", "*.csv")])
            if not csv_path:
                print("CSVファイルが選択されませんでした。")
                return
                
            # CSVからデータを読み込む
            self.load_csv(csv_path)
        
        # 最初のページを読み込む
        if self.page_product_names:
            self.current_page_index = 0
            self.current_page = self.page_product_names[0]
            self.load_page(self.current_page)
            self.root.mainloop()
        else:
            print("ページ商品名が見つかりませんでした。")
    
    def process_page_data(self):
        # create_all.pyから渡されたデータを処理
        for page_name, cards in self.page_data:
            self.page_product_names.append(page_name)
            for card in cards:
                self.all_cards.append({
                    'item_code': card['item_code'],
                    'product_name': card['product_name'],
                    'color_jp': card['color'],
                    'page_product_name': page_name
                })
                # プレビュー用画像パスを出力（デバッグ用）
                if len(self.all_cards) <= 3:  # 最初の3枚だけ表示
                    size = 'full' if len(self.all_cards) == 1 else 'half'
                    print(f"プレビュー用画像パス: output\\{card['item_code']}-{card['product_name']}_{card['color']}_{size}.png")
    
    def load_csv(self, csv_path):
        # CSVからデータを読み込む
        self.all_cards = []
        self.page_product_names = set()
        
        with open(csv_path, 'r', encoding='shift_jis') as f:
            reader = csv.reader(f)
            next(reader)  # ヘッダーをスキップ
            
            for row in reader:
                if len(row) < 6:  # F列まで必要
                    print(f"スキップ: 列数が足りません")
                    continue
                
                item_code = row[0]      # A列：商品コード
                product_name = row[2]    # C列：商品名
                color_jp = row[3]        # D列：カラー名（日本語）
                page_product_name = row[5]  # F列：ページ商品名
                
                self.all_cards.append({
                    'item_code': item_code,
                    'product_name': product_name,
                    'color_jp': color_jp,
                    'page_product_name': page_product_name
                })
                self.page_product_names.add(page_product_name)
        
        # ページ商品名をソート
        self.page_product_names = sorted(list(self.page_product_names))

class PageCreator:
    def __init__(self):
        self.page_generator = A4PageGenerator()

    def process_csv_data(self, data_rows, output_dir, a4_output_dir):
        # データをCardPlacementInterfaceに渡すための形式に変換
        page_data = []
        product_groups = {}
        
        # 商品グループごとにデータを整理
        for row in data_rows:
            if len(row) >= 6:  # 必要な列数をチェック
                page_product_name = row[5]  # F列：ページ商品名
                if page_product_name not in product_groups:
                    product_groups[page_product_name] = []
                
                card_data = {
                    'item_code': row[0],
                    'product_name': row[2],
                    'color': row[3],
                }
                product_groups[page_product_name].append(card_data)

        # ページデータを作成
        for page_name, cards in product_groups.items():
            page_data.append((page_name, cards))

        # CardPlacementInterfaceを使用してページを生成
        interface = CardPlacementInterface(page_data=page_data)
        interface.main()

    def process_csv(self):
        csv_path = filedialog.askopenfilename(title="CSVファイルを選択", filetypes=[("CSV files", "*.csv")])
        if not csv_path:
            print("CSVファイルが選択されませんでした。")
            return []

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

        # CardPlacementInterfaceを使用してページを生成
        interface = CardPlacementInterface()
        interface.main()

        # 処理結果を返す（空のリストでも返す）
        return rows + data_rows

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    return result['encoding']

def main():
    interface = CardPlacementInterface()
    interface.main()

if __name__ == "__main__":
    main()
import dataclasses
from typing import List, Tuple, Dict

@dataclasses.dataclass
class CardLayout:
    name: str  # パターン名
    description: str  # パターンの説明
    positions: List[Tuple[int, int]]  # 各カードの位置 [(x, y), ...]
    sizes: List[str]  # 各カードのサイズ ['full' or 'half', ...]

class CardLayoutManager:
    def __init__(self):
        # A4サイズ（300dpi、横）
        self.a4_width = 3508
        self.a4_height = 2480
        
        # カードサイズ定義（300dpi）
        self.full_width = 1704  # フルサイズの幅
        self.full_height = 2000  # フルサイズの高さ
        self.half_width = 852   # ハーフサイズの幅
        self.half_height = 1000 # ハーフサイズの高さ
        
        # 余白
        self.margin_top = 400

    def get_layouts_for_count(self, num_cards: int) -> Dict[str, CardLayout]:
        """指定されたカード枚数に対応する配置パターンを返す"""
        if num_cards == 4:
            return {
                "pattern1": self._get_layout_4cards_pattern1(),
                "pattern2": self._get_layout_4cards_pattern2(),
            }
        elif num_cards == 3:
            return {
                "pattern1": self._get_layout_3cards(),
            }
        elif num_cards == 2:
            return {
                "pattern1": self._get_layout_2cards(),
            }
        elif num_cards == 1:
            return {
                "pattern1": self._get_layout_1card(),
            }
        elif num_cards == 5:
            return {
                "pattern1": self._get_layout_5cards(),
            }
        elif num_cards == 6:
            return {
                "pattern1": self._get_layout_6cards(),
            }
        elif num_cards == 7:
            return {
                "pattern1": self._get_layout_7cards(),
            }
        elif num_cards == 8:
            return {
                "pattern1": self._get_layout_8cards(),
            }
        return {}

    def _get_layout_4cards_pattern1(self) -> CardLayout:
        """4枚パターン1: フル1枚 + ハーフ3枚"""
        # フル画像1枚 + 残りをハーフサイズで配置
        num_half = 3
        num_half_cols = 2  # ハーフサイズ画像の列数

        # 全体の幅を計算（フル幅 + ハーフ幅 * 列数）
        total_width = self.full_width + self.half_width * num_half_cols
        # 中央寄せのためのオフセットを計算
        center_offset = int((self.a4_width - total_width) / 2)

        # フル画像の位置（左側）
        positions = [(center_offset, self.margin_top)]
        sizes = ['full']

        # ハーフ画像の位置（右側に格子状に配置）
        half_start_x = center_offset + self.full_width
        for i in range(num_half):
            col = i % num_half_cols
            row = i // num_half_cols
            x = half_start_x + col * self.half_width
            y = self.margin_top + row * self.half_height
            positions.append((x, y))
            sizes.append('half')

        return CardLayout(
            name="パターン1",
            description="フルサイズ1枚 + ハーフサイズ3枚（右側格子状）",
            positions=positions,
            sizes=sizes
        )

    def _get_layout_4cards_pattern2(self) -> CardLayout:
        """4枚パターン2: ハーフ4枚を2x2で配置"""
        num_cols = 2
        num_rows = 2
        total_width = self.half_width * num_cols
        center_offset_x = int((self.a4_width - total_width) / 2)

        positions = []
        sizes = []
        for i in range(4):
            col = i % num_cols
            row = i // num_cols
            x = center_offset_x + col * self.half_width
            y = self.margin_top + row * self.half_height
            positions.append((x, y))
            sizes.append('half')

        return CardLayout(
            name="パターン2",
            description="ハーフサイズ4枚を2×2で配置（中央）",
            positions=positions,
            sizes=sizes
        )

    def _get_layout_3cards(self) -> CardLayout:
        """3枚: フル1枚 + ハーフ2枚"""
        # 全体の幅を計算（フル幅 + ハーフ幅）
        total_width = self.full_width + self.half_width
        # 中央寄せのためのオフセットを計算
        center_offset = int((self.a4_width - total_width) / 2)
        
        # フル画像とハーフ画像の位置を計算
        x1 = center_offset  # フル画像の開始位置
        x2 = x1 + self.full_width  # ハーフ画像の開始位置（フル画像の直後）
        
        positions = [
            (x1, self.margin_top),  # 左（フル）
            (x2, self.margin_top),  # 右上（ハーフ）
            (x2, self.margin_top + self.half_height)  # 右下（ハーフ）
        ]
        sizes = ['full', 'half', 'half']

        return CardLayout(
            name="パターン1",
            description="フルサイズ1枚 + ハーフサイズ2枚（右側縦並び）",
            positions=positions,
            sizes=sizes
        )

    def _get_layout_2cards(self) -> CardLayout:
        """2枚: フル2枚を横に並べて中央揃え"""
        total_width = self.full_width * 2
        center_offset = int((self.a4_width - total_width) / 2)
        positions = [
            (center_offset, self.margin_top),
            (center_offset + self.full_width, self.margin_top)
        ]
        sizes = ['full', 'full']

        return CardLayout(
            name="パターン1",
            description="フルサイズ2枚を横並び",
            positions=positions,
            sizes=sizes
        )

    def _get_layout_1card(self) -> CardLayout:
        """1枚: フル1枚を中央配置"""
        positions = [(int((self.a4_width - self.full_width) / 2), self.margin_top)]
        sizes = ['full']

        return CardLayout(
            name="パターン1",
            description="フルサイズ1枚を中央配置",
            positions=positions,
            sizes=sizes
        )

    def _get_layout_5cards(self) -> CardLayout:
        """5枚: フル1枚 + ハーフ4枚"""
        # フル画像1枚 + 残りをハーフサイズで配置
        num_half = 4
        num_half_cols = 2  # ハーフサイズ画像の列数

        # 全体の幅を計算（フル幅 + ハーフ幅 * 列数）
        total_width = self.full_width + self.half_width * num_half_cols
        # 中央寄せのためのオフセットを計算
        center_offset = int((self.a4_width - total_width) / 2)

        # フル画像の位置（左側）
        positions = [(center_offset, self.margin_top)]
        sizes = ['full']

        # ハーフ画像の位置（右側に格子状に配置）
        half_start_x = center_offset + self.full_width
        for i in range(num_half):
            col = i % num_half_cols
            row = i // num_half_cols
            x = half_start_x + col * self.half_width
            y = self.margin_top + row * self.half_height
            positions.append((x, y))
            sizes.append('half')

        return CardLayout(
            name="パターン1",
            description="フルサイズ1枚 + ハーフサイズ4枚（右側格子状）",
            positions=positions,
            sizes=sizes
        )

    def _get_layout_6cards(self) -> CardLayout:
        """6枚: ハーフ6枚を3x2で配置"""
        num_cols = 3
        num_rows = 2
        total_width = self.half_width * num_cols
        center_offset_x = int((self.a4_width - total_width) / 2)

        positions = []
        sizes = []
        for i in range(6):
            col = i % num_cols
            row = i // num_cols
            x = center_offset_x + col * self.half_width
            y = self.margin_top + row * self.half_height
            positions.append((x, y))
            sizes.append('half')

        return CardLayout(
            name="パターン1",
            description="ハーフサイズ6枚を3×2で配置",
            positions=positions,
            sizes=sizes
        )

    def _get_layout_7cards(self) -> CardLayout:
        """7枚: ハーフ7枚を4+3で配置"""
        # 上段4枚
        total_width = self.half_width * 4
        center_offset_x = int((self.a4_width - total_width) / 2)

        positions = []
        sizes = []

        # 上段4枚
        for i in range(4):
            x = center_offset_x + i * self.half_width
            y = self.margin_top
            positions.append((x, y))
            sizes.append('half')

        # 下段3枚
        total_width = self.half_width * 3
        center_offset_x = int((self.a4_width - total_width) / 2)
        for i in range(3):
            x = center_offset_x + i * self.half_width
            y = self.margin_top + self.half_height
            positions.append((x, y))
            sizes.append('half')

        return CardLayout(
            name="パターン1",
            description="ハーフサイズ7枚を4+3で配置",
            positions=positions,
            sizes=sizes
        )

    def _get_layout_8cards(self) -> CardLayout:
        """8枚: ハーフ8枚を4x2で配置"""
        num_cols = 4
        num_rows = 2
        total_width = self.half_width * num_cols
        center_offset_x = int((self.a4_width - total_width) / 2)

        positions = []
        sizes = []
        for i in range(8):
            col = i % num_cols
            row = i // num_cols
            x = center_offset_x + col * self.half_width
            y = self.margin_top + row * self.half_height
            positions.append((x, y))
            sizes.append('half')

        return CardLayout(
            name="パターン1",
            description="ハーフサイズ8枚を4×2で配置",
            positions=positions,
            sizes=sizes
        ) 
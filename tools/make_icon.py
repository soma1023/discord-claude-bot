# -*- coding: utf-8 -*-
"""デスクトップ用のアイコン（.ico）を作る。

図柄は画面の中身と揃えて、コメント密度のグラフ。一番高い棒だけ色を変え、
「跳ねた瞬間を見つける道具」だと小さい表示でも分かるようにしている。

サイズごとに描き分けているのは、16pxを大きい絵の縮小で作ると棒がぼやけて
潰れるため。棒は整数座標に合わせ、角丸だけを多重標本でなめらかにする。

外部ライブラリは使わない（PNGもICOも自前で組む）。アイコンを作るためだけに
導入物を増やすと、手元で動かすときの失敗箇所が増える。

    python tools/make_icon.py
"""

import os
import struct
import zlib

SIZES = [16, 20, 24, 32, 48, 64, 128, 256]
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATIC = os.path.join(HERE, "stream_highlight", "static")
OUT = os.path.join(STATIC, "icon.ico")

BG = (22, 27, 34)               # --panel #161b22
BAR = (88, 166, 255)            # --accent #58a6ff
PEAK = (126, 231, 135)          # --good  #7ee787

# 16単位の方眼で図柄を決め、各サイズへ拡大する
GRID = 16.0
BARS = [3.5, 9.0, 5.5, 4.2]     # 棒の高さ（2本目が見せ場）
PEAK_INDEX = 1
BAR_W = 2.0
BAR_GAP = 1.3
LEFT = 2.0
BASELINE = 13.0                 # 棒の下端（上からの距離）
CORNER = 0.22                   # 角丸の半径（一辺に対する比）
SUB = 4                         # 角丸の多重標本


def _inside_rounded(x, y, side, radius):
    """角丸四角の内側かどうか。"""
    near_x = min(max(x, radius), side - radius)
    near_y = min(max(y, radius), side - radius)
    dx = x - near_x
    dy = y - near_y
    return dx * dx + dy * dy <= radius * radius


def draw_icon(size):
    """RGBAの画素列を返す。"""
    radius = size * CORNER
    pixels = bytearray(size * size * 4)

    for py in range(size):
        for px in range(size):
            # 角の内外は多重標本で求め、縁のがたつきを抑える
            hits = 0
            for sy in range(SUB):
                for sx in range(SUB):
                    if _inside_rounded(px + (sx + 0.5) / SUB,
                                       py + (sy + 0.5) / SUB, size, radius):
                        hits += 1
            if hits:
                at = (py * size + px) * 4
                pixels[at:at + 3] = bytes(BG)
                pixels[at + 3] = round(255 * hits / (SUB * SUB))

    unit = size / GRID
    for index, height in enumerate(BARS):
        left = LEFT + index * (BAR_W + BAR_GAP)
        # 棒は整数座標へ丸める。半端な位置だと小さいサイズでぼやける。
        x0 = int(round(left * unit))
        x1 = max(x0 + 1, int(round((left + BAR_W) * unit)))
        y1 = int(round(BASELINE * unit))
        y0 = min(y1 - 1, int(round((BASELINE - height) * unit)))
        color = bytes(PEAK if index == PEAK_INDEX else BAR) + b"\xff"
        for py in range(max(0, y0), min(size, y1)):
            for px in range(max(0, x0), min(size, x1)):
                at = (py * size + px) * 4
                pixels[at:at + 4] = color
    return pixels


def to_png(pixels, size):
    """RGBAの画素列をPNGにする。"""
    raw = bytearray()
    for row in range(size):
        raw.append(0)                         # フィルタなし
        raw += pixels[row * size * 4:(row + 1) * size * 4]

    def chunk(kind, body):
        return (struct.pack(">I", len(body)) + kind + body
                + struct.pack(">I", zlib.crc32(kind + body) & 0xFFFFFFFF))

    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
            + chunk(b"IEND", b""))


def write_ico(blobs, sizes, path):
    """PNGを並べた .ico を組む。256px は仕様上サイズ欄が 0。"""
    offset = 6 + 16 * len(blobs)
    out = struct.pack("<HHH", 0, 1, len(blobs))
    for size, blob in zip(sizes, blobs):
        side = 0 if size >= 256 else size
        out += struct.pack("<BBBBHHII", side, side, 0, 0, 1, 32,
                           len(blob), offset)
        offset += len(blob)
    with open(path, "wb") as fh:
        fh.write(out + b"".join(blobs))


def main():
    blobs = []
    for size in SIZES:
        blobs.append(to_png(draw_icon(size), size))
    write_ico(blobs, SIZES, OUT)
    print("%s / %d バイト / %s" % (OUT, os.path.getsize(OUT),
                                   " ".join("%d" % s for s in SIZES)))

    # 目で確かめられるよう、一番大きいものをPNGでも残す
    with open(os.path.join(STATIC, "icon.png"), "wb") as fh:
        fh.write(blobs[-1])


if __name__ == "__main__":
    main()

import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, portrait
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
import os
import re
import unicodedata

# ==========================================
# 定数・設定
# ==========================================

# ★ 太字設定 ★
# Trueにすると、文字の輪郭をなぞって太く見せます（擬似ボールド）
# MS明朝に切り替わった文字もこれで太くなります。
USE_BOLD = True

# 文字サイズに対する線の太さの比率
# 0.03あたりが程よい太さです。太すぎる場合は0.01〜0.02に下げてください。
BOLD_WIDTH_RATIO = 0.05

# --- フォント設定 ---
# HGRGEで表示できない文字と、代替フォントの対応表
SPECIAL_CHARS_FONT_MAP = {
    '嵗': 'MSMincho',
    '俻': 'MSMincho',
    # 必要に応じて追加してください
    # '𠮷': 'MSMincho', 
}
DEFAULT_FONT = "MSMincho"

LEADING_PROHIBITED_CHARS = {
    '。', '、', '」', '』', ')', '）', ']', '}'
}

# --- 文字位置調整設定 ---

HANGING_PUNCTUATION_ADJUSTMENTS = {
    '。': {'x_offset': 8, 'y_offset': -2},
    '、': {'x_offset': 8, 'y_offset': -2},
    '」': {'x_offset': 3, 'y_offset': -5},
    '』': {'x_offset': 2, 'y_offset': -5},
    ')': {'x_offset': 2, 'y_offset': -5},
    '）': {'x_offset': 4, 'y_offset': 0},
    ']': {'x_offset': 4, 'y_offset': 0},
    '}': {'x_offset': 4, 'y_offset': 0},
}

PUNCTUATION_ADJUSTMENTS = {
    '。': {'x_offset': 7, 'y_offset': 5},
    '、': {'x_offset': 7, 'y_offset': 5},
    'っ': {'x_offset': 1, 'y_offset': 0},
    'ゃ': {'x_offset': 1, 'y_offset': 0},
    'ゅ': {'x_offset': 1, 'y_offset': 0},
    'ょ': {'x_offset': 1, 'y_offset': 0},
    'ぁ': {'x_offset': 1, 'y_offset': 0},
    'ぃ': {'x_offset': 1, 'y_offset': 0},
    'ぅ': {'x_offset': 1, 'y_offset': 0},
    'ぇ': {'x_offset': 1, 'y_offset': 0},
    'ぉ': {'x_offset': 1, 'y_offset': 0},
    'ー': {'x_offset': 4, 'y_offset': 8},
    '(': {'x_offset': 4.5, 'y_offset': 10},
    ')': {'x_offset': 4.5, 'y_offset': 3},
    '（': {'x_offset': 4, 'y_offset': 12.5},
    '）': {'x_offset': 4, 'y_offset': 7.5},
    '[': {'x_offset': 3, 'y_offset': -5},
    ']': {'x_offset': 3, 'y_offset': -5},
    '{': {'x_offset': 3, 'y_offset': -5},
    '}': {'x_offset': 3, 'y_offset': -5},
    '「': {'x_offset': 5, 'y_offset': 12},
    '」': {'x_offset': 3, 'y_offset': 4},
    '『': {'x_offset': 3, 'y_offset': -5},
    '』': {'x_offset': 3, 'y_offset': -5},
    '-': {'x_offset': 12, 'y_offset': 6},
    '→': {'x_offset': 11.5, 'y_offset': 8},
    '←': {'x_offset': 11.5, 'y_offset': 8},
    '↑': {'x_offset': 11.5, 'y_offset': 8},
    '↓': {'x_offset': 11.5, 'y_offset': 8},
    'a': {'x_offset': 4, 'y_offset': 7},
    'b': {'x_offset': 4, 'y_offset': 7},
    'c': {'x_offset': 4, 'y_offset': 7},
    'd': {'x_offset': 4, 'y_offset': 7},
    'e': {'x_offset': 4, 'y_offset': 7},
    'f': {'x_offset': 4, 'y_offset': 7},
    'g': {'x_offset': 4, 'y_offset': 7},
    'h': {'x_offset': 4, 'y_offset': 7},
    'i': {'x_offset': 4, 'y_offset': 7},
    'j': {'x_offset': 4, 'y_offset': 7},
    'k': {'x_offset': 4, 'y_offset': 7},
    'l': {'x_offset': 4, 'y_offset': 7},
    'm': {'x_offset': 4, 'y_offset': 7},
    'n': {'x_offset': 4, 'y_offset': 7},
    'o': {'x_offset': 4, 'y_offset': 7},
    'p': {'x_offset': 4, 'y_offset': 7},
    'q': {'x_offset': 4, 'y_offset': 7},
    'r': {'x_offset': 4, 'y_offset': 7},
    's': {'x_offset': 4, 'y_offset': 7},
    't': {'x_offset': 4, 'y_offset': 7},
    'u': {'x_offset': 4, 'y_offset': 7},
    'v': {'x_offset': 4, 'y_offset': 7},
    'w': {'x_offset': 4, 'y_offset': 7},
    'x': {'x_offset': 4, 'y_offset': 7},
    'y': {'x_offset': 4, 'y_offset': 7},
    'z': {'x_offset': 4, 'y_offset': 7},
    'A': {'x_offset': 4, 'y_offset': 7},
    'B': {'x_offset': 4, 'y_offset': 7},
    'C': {'x_offset': 4, 'y_offset': 7},
    'D': {'x_offset': 4, 'y_offset': 7},
    'E': {'x_offset': 4, 'y_offset': 7},
    'F': {'x_offset': 4, 'y_offset': 7},
    'G': {'x_offset': 4, 'y_offset': 7},
    'H': {'x_offset': 4, 'y_offset': 7},
    'I': {'x_offset': 4, 'y_offset': 7},
    'J': {'x_offset': 4, 'y_offset': 7},
    'K': {'x_offset': 4, 'y_offset': 7},
    'L': {'x_offset': 4, 'y_offset': 7},
    'M': {'x_offset': 4, 'y_offset': 7},
    'N': {'x_offset': 4, 'y_offset': 7},
    'O': {'x_offset': 4, 'y_offset': 7},
    'P': {'x_offset': 4, 'y_offset': 7},
    'Q': {'x_offset': 4, 'y_offset': 7},
    'R': {'x_offset': 4, 'y_offset': 7},
    'S': {'x_offset': 4, 'y_offset': 7},
    'T': {'x_offset': 4, 'y_offset': 7},
    'U': {'x_offset': 4, 'y_offset': 7},
    'V': {'x_offset': 4, 'y_offset': 7},
    'W': {'x_offset': 4, 'y_offset': 7},
    'X': {'x_offset': 4, 'y_offset': 7},
    'Y': {'x_offset': 4, 'y_offset': 7},
    'Z': {'x_offset': 4, 'y_offset': 7},
    '0': {'x_offset': 2.5, 'y_offset': 0},
    '1': {'x_offset': 2.5, 'y_offset': 0},
    '2': {'x_offset': 2.5, 'y_offset': 0},
    '3': {'x_offset': 2.5, 'y_offset': 0},
    '4': {'x_offset': 2.5, 'y_offset': 0},
    '5': {'x_offset': 2.5, 'y_offset': 0},
    '6': {'x_offset': 2.5, 'y_offset': 0},
    '7': {'x_offset': 2.5, 'y_offset': 0},
    '8': {'x_offset': 2.5, 'y_offset': 0},
    '9': {'x_offset': 2.5, 'y_offset': 0},
    '!': {'x_offset': 2.5, 'y_offset': 0},
    'Ａ': {'x_offset': 3.5, 'y_offset':9},
    'Ｂ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｃ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｄ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｅ': {'x_offset': 3.5, 'y_offset': 9}, 
    'Ｆ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｇ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｈ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｉ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｊ': {'x_offset': 3.5, 'y_offset': 9}, 
    'Ｋ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｌ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｍ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｎ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｏ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｐ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｑ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｒ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｓ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｔ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｕ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｖ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｗ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｘ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｙ': {'x_offset': 3.5, 'y_offset': 9},
    'Ｚ': {'x_offset': 3.5, 'y_offset': 9},
    'ａ': {'x_offset': 3.5, 'y_offset': 9},
    'ｂ': {'x_offset': 3.5, 'y_offset': 9},
    'ｃ': {'x_offset': 3.5, 'y_offset': 9}, 
    'ｄ': {'x_offset': 3.5, 'y_offset': 9},
    'ｅ': {'x_offset': 3.5, 'y_offset': 9},
    'ｆ': {'x_offset': 3.5, 'y_offset': 9},
    'ｇ': {'x_offset': 3.5, 'y_offset': 9},
    'ｈ': {'x_offset': 3.5, 'y_offset': 9},
    'ｉ': {'x_offset': 3.5, 'y_offset': 9},
    'ｊ': {'x_offset': 3.5, 'y_offset': 9},
    'ｋ': {'x_offset': 3.5, 'y_offset': 9},
    'ｌ': {'x_offset': 3.5, 'y_offset': 9},
    'ｍ': {'x_offset': 3.5, 'y_offset': 9},
    'ｎ': {'x_offset': 3.5, 'y_offset': 9},
    'ｏ': {'x_offset': 3.5, 'y_offset': 9},
    'ｐ': {'x_offset': 3.5, 'y_offset': 9},
    'ｑ': {'x_offset': 3.5, 'y_offset': 9},
    'ｒ': {'x_offset': 3.5, 'y_offset': 9},
    'ｓ': {'x_offset': 3.5, 'y_offset': 9},
    'ｔ': {'x_offset': 3.5, 'y_offset': 9},
    'ｕ': {'x_offset': 3.5, 'y_offset': 9},
    'ｖ': {'x_offset': 3.5, 'y_offset': 9},
    'ｗ': {'x_offset': 3.5, 'y_offset': 9},
    'ｘ': {'x_offset': 3.5, 'y_offset': 9},
    'ｙ': {'x_offset': 3.5, 'y_offset': 9},
    'ｚ': {'x_offset': 3.5, 'y_offset': 9},
    '，': {'x_offset': 7, 'y_offset':5},

}

WORK_NAME_ADJUSTMENTS = {
    '。': {'x_offset': 5, 'y_offset': 3},
    '、': {'x_offset': 5, 'y_offset': 3},
    'っ': {'x_offset': 1, 'y_offset': 0},
    'ゃ': {'x_offset': 1, 'y_offset': 0},
    'ゅ': {'x_offset': 1, 'y_offset': 0},
    'ょ': {'x_offset': 1, 'y_offset': 0},
    'ぁ': {'x_offset': 1, 'y_offset': 0},
    'ぃ': {'x_offset': 1, 'y_offset': 0},
    'ぅ': {'x_offset': 1, 'y_offset': 0},
    'ぇ': {'x_offset': 1, 'y_offset': 0},
    'ぉ': {'x_offset': 0.5, 'y_offset': 0},
    'ー': {'x_offset': 7, 'y_offset': 15},
    '(': {'x_offset': 9.5, 'y_offset': 8},
    ')': {'x_offset': 9.5, 'y_offset': 2},
    '（': {'x_offset': 9.5, 'y_offset': 10},
    '）': {'x_offset': 9.5, 'y_offset': 6},
    '[': {'x_offset': 2, 'y_offset': -3},
    ']': {'x_offset': 2, 'y_offset': -3},
    '{': {'x_offset': 2, 'y_offset': -3},
    '}': {'x_offset': 2, 'y_offset': -3},
    '「': {'x_offset': 0, 'y_offset': 0},
    '」': {'x_offset': 9, 'y_offset': 4},
    '『': {'x_offset': 2, 'y_offset': -3},
    '』': {'x_offset': 2, 'y_offset': -3},
    '-': {'x_offset': 10, 'y_offset': 4},
    '→': {'x_offset': 9.5, 'y_offset': 6},
    '←': {'x_offset': 9.5, 'y_offset': 6},
    '↑': {'x_offset': 9.5, 'y_offset': 6},
    '↓': {'x_offset': 9.5, 'y_offset': 6},
    'A': {'x_offset': 5, 'y_offset': 11},
    'B': {'x_offset': 5, 'y_offset': 11},
    'C': {'x_offset': 5, 'y_offset': 11},
    'D': {'x_offset': 5, 'y_offset': 11},
    'E': {'x_offset': 5, 'y_offset': 11},
    'F': {'x_offset': 5, 'y_offset': 11},
    'G': {'x_offset': 5, 'y_offset': 11},
    'H': {'x_offset': 5, 'y_offset': 11},
    'I': {'x_offset': 5, 'y_offset': 11},
    'J': {'x_offset': 5, 'y_offset': 11},
    'K': {'x_offset': 5, 'y_offset': 11},
    'L': {'x_offset': 5, 'y_offset': 11},
    'M': {'x_offset': 5, 'y_offset': 11},
    'N': {'x_offset': 5, 'y_offset': 11},
    'O': {'x_offset': 5, 'y_offset': 11},
    'P': {'x_offset': 5, 'y_offset': 11},
    'Q': {'x_offset': 5, 'y_offset': 11},
    'R': {'x_offset': 5, 'y_offset': 11},
    'S': {'x_offset': 5, 'y_offset': 11},
    'T': {'x_offset': 5, 'y_offset': 11},
    'U': {'x_offset': 5, 'y_offset': 11},
    'V': {'x_offset': 5, 'y_offset': 11},
    'W': {'x_offset': 5, 'y_offset': 11},
    'X': {'x_offset': 5, 'y_offset': 11},
    'Y': {'x_offset': 5, 'y_offset': 11},
    'Z': {'x_offset': 5, 'y_offset': 11},
    'a': {'x_offset': 5, 'y_offset': 11},
    'b': {'x_offset': 5, 'y_offset': 11},
    'c': {'x_offset': 5, 'y_offset': 11},
    'd': {'x_offset': 5, 'y_offset': 11},
    'e': {'x_offset': 5, 'y_offset': 11},
    'f': {'x_offset': 5, 'y_offset': 11},
    'g': {'x_offset': 5, 'y_offset': 11},
    'h': {'x_offset': 5, 'y_offset': 11},
    'i': {'x_offset': 5, 'y_offset': 11},
    'j': {'x_offset': 5, 'y_offset': 11},
    'k': {'x_offset': 5, 'y_offset': 11},
    'l': {'x_offset': 5, 'y_offset': 11},
    'm': {'x_offset': 5, 'y_offset': 11},
    'n': {'x_offset': 5, 'y_offset': 11},
    'o': {'x_offset': 5, 'y_offset': 11},
    'p': {'x_offset': 5, 'y_offset': 11},
    'q': {'x_offset': 5, 'y_offset': 11},
    'r': {'x_offset': 5, 'y_offset': 11},
    's': {'x_offset': 5, 'y_offset': 11},
    't': {'x_offset': 5, 'y_offset': 11},
    'u': {'x_offset': 5, 'y_offset': 11},
    'v': {'x_offset': 5, 'y_offset': 11},
    'w': {'x_offset': 5, 'y_offset': 11},
    'x': {'x_offset': 5, 'y_offset': 11},
    'y': {'x_offset': 5, 'y_offset': 11},
    'z': {'x_offset': 5, 'y_offset': 11},
    '〈': {'x_offset': 7.5, 'y_offset':15},
    '〉': {'x_offset': 7.5, 'y_offset':15},
    '～': {'x_offset': 7.5, 'y_offset':15},
    '!': {'x_offset': 6, 'y_offset':15},
    '，': {'x_offset': 12, 'y_offset':5},
}

WORK_INFO_ADJUSTMENTS = WORK_NAME_ADJUSTMENTS.copy()
WORK_INFO_ADJUSTMENTS.update({
    '「': {'x_offset': 10, 'y_offset': 14, 'angle': 270},
    '」': {'x_offset': 6, 'y_offset': 15, 'angle': 270},
    
})

NAME_ADJUSTMENTS = {
    '。': {'x_offset': 4, 'y_offset': 2},
    '、': {'x_offset': 4, 'y_offset': 2},
    'っ': {'x_offset': 1, 'y_offset': 1},
    'ゃ': {'x_offset': 1, 'y_offset': 1},
    'ゅ': {'x_offset': 1, 'y_offset': 1},
    'ょ': {'x_offset': 1, 'y_offset': 1},
    'ぁ': {'x_offset': 1, 'y_offset': 1},
    'ぃ': {'x_offset': 1, 'y_offset': 1},
    'ぅ': {'x_offset': 1, 'y_offset': 1},
    'ぇ': {'x_offset': 1, 'y_offset': 1},
    'ぉ': {'x_offset': 1, 'y_offset': 1},
    'ー': {'x_offset': 8, 'y_offset': 5},
    '(': {'x_offset': 8.5, 'y_offset': 7},
    ')': {'x_offset': 8.5, 'y_offset': 1},
    '（': {'x_offset': 5.5, 'y_offset': 10},
    '）': {'x_offset': 5, 'y_offset': 11},
    '　': {'x_offset': 0, 'y_offset': 0},
}

DEPARTMENT_YEAR_ADJUSTMENTS = {
    '。': {'x_offset': 6, 'y_offset': 4},
    '、': {'x_offset': 6, 'y_offset': 4},
    '0': {'x_offset': 1.5, 'y_offset': 0},
    '1': {'x_offset': 1.5, 'y_offset': 0},
    '2': {'x_offset': 1.5, 'y_offset': 0},
    '3': {'x_offset': 1.5, 'y_offset': 0},
    '4': {'x_offset': 1.5, 'y_offset': 0},
    '5': {'x_offset': 1.5, 'y_offset': 0},
    '6': {'x_offset': 1.5, 'y_offset': 0},
    '7': {'x_offset': 1.5, 'y_offset': 0},
    '8': {'x_offset': 1.5, 'y_offset': 0},
    '9': {'x_offset': 1.5, 'y_offset': 0},
    '学': {'x_offset': 0, 'y_offset': 0},
    '部': {'x_offset': 0, 'y_offset': 0},
    '年': {'x_offset': 0, 'y_offset': 0},
    '生': {'x_offset': 0, 'y_offset': 0},
}

ROTATED_CHARS = {'(', ')', '（', '）', '[', ']', '{', '}', '「', '」', '『', '』', 'ー','-','→','←','↑','↓','＜','＞','〈','〉','～','!',
                'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z',
                'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
                'Ａ','Ｂ','Ｃ','Ｄ','Ｅ','Ｆ','Ｇ','Ｈ','Ｉ','Ｊ','Ｋ','Ｌ','Ｍ','Ｎ','Ｏ','Ｐ','Ｑ','Ｒ','Ｓ','Ｔ','Ｕ','Ｖ','Ｗ','Ｘ','Ｙ','Ｚ',
                'ａ','ｂ','ｃ','ｄ','ｅ','ｆ','ｇ','ｈ','ｉ','ｊ','ｋ','ｌ','ｍ','ｎ','ｏ','ｐ','ｑ','ｒ','ｓ','ｔ','ｕ','ｖ','ｗ','ｘ','ｙ','ｚ'}

COORDINATES = {
    '学部学年': {
        'x': 250, 'y': 780, 'font_size': 18, 'char_spacing': 1.0,
        'wrap': False, 'horizontal': False,
        'adjustments': DEPARTMENT_YEAR_ADJUSTMENTS
    },
    '名前': {
        'handler': 'name_and_furigana',
        'x': 250,
        'y': 495,
        'name_font_size': 18,
        'furigana_font_size': 13,
        'furigana_x_offset': 2.75,
        'furigana_y_spacing': 0.5,
        'char_spacing': 1.0,
        'adjustments': NAME_ADJUSTMENTS
    },
    '作品情報': {
        'x': 206, 'y': 780, 'font_size': 18, 'char_spacing': 1.0,
        'wrap': False, 'horizontal': False,
        'adjustments': WORK_INFO_ADJUSTMENTS
    },
    '釈文': {
        'x': [152, 162, 172],
        'y': 786, 
        'font_size': 10, 'char_spacing': 1.0, 'wrap': True,
        'max_chars': 55, 'line_spacing': 20, 'horizontal': False,
        'adjustments': PUNCTUATION_ADJUSTMENTS
    },
    'コメント': {
        'x': [70, 80, 85, 95, 105],
        'y': 786,
        'font_size': 10, 'char_spacing': 1.0, 'wrap': True,
        'max_chars': 55, 'line_spacing': 20, 'horizontal': False,
        'adjustments': PUNCTUATION_ADJUSTMENTS
    },
    '臨書解説': {
        'x': 25, 'y': 170, 'font_size': 10, 'char_spacing': 1,
        'wrap': True, 'max_chars': 25, 'line_spacing': 30, 'horizontal': True,
        'adjustments': PUNCTUATION_ADJUSTMENTS
    },
    '作品情報（法帖解説）': {
        'x': 150, 'y': 190, 'font_size': 14, 'char_spacing': 1,
        'wrap': True, 'max_chars': 25, 'line_spacing': 30, 'horizontal': True, 'centered': True,
        'adjustments': WORK_NAME_ADJUSTMENTS
    },
    '再提出': {
        'x': 150, 'y': 5, 'font_size': 14, 'char_spacing': 1,
        'wrap': True, 'max_chars': 25, 'line_spacing': 30, 'horizontal': True, 'centered': True,
        'adjustments': PUNCTUATION_ADJUSTMENTS
    },
}

OFFSET_X = 300

# ==========================================
# 描画関数群 (太字対応済み)
# ==========================================

def draw_name_and_furigana(canvas, name, furigana, x, y, name_font_size, furigana_font_size, char_spacing, adjustments, furigana_x_offset, furigana_y_spacing):
    if adjustments is None:
        adjustments = {}
    current_y = y
    
    # --- 名前描画 ---
    for char in str(name):
        target_font = SPECIAL_CHARS_FONT_MAP.get(char, DEFAULT_FONT)
        canvas.setFont(target_font, name_font_size)
        
        adj = adjustments.get(char, {'x_offset': 0, 'y_offset': 0})
        draw_x = x + adj.get('x_offset', 0)
        draw_y = current_y + adj.get('y_offset', 0)
        
        # ★太字処理 (2 Tr = Fill & Stroke)
        if USE_BOLD:
            canvas.setLineWidth(name_font_size * BOLD_WIDTH_RATIO)
            canvas.setStrokeColorRGB(0, 0, 0)
            canvas._code.append('2 Tr')
        
        if char in ROTATED_CHARS:
            canvas.saveState()
            canvas.translate(draw_x, draw_y)
            canvas.rotate(270)
            canvas.drawString(0, -name_font_size / 4, char)
            canvas.restoreState()
        else:
            canvas.drawString(draw_x, draw_y, char)
            
        # ★太字リセット (0 Tr = Fill Only)
        if USE_BOLD:
            canvas._code.append('0 Tr')

        current_y -= name_font_size * char_spacing

    if name and furigana:
        current_y -= furigana_y_spacing

    furigana_base_x = x + furigana_x_offset
    
    # --- ふりがな描画 ---
    for char in str(furigana):
        target_font = SPECIAL_CHARS_FONT_MAP.get(char, DEFAULT_FONT)
        canvas.setFont(target_font, furigana_font_size)

        adj = adjustments.get(char, {'x_offset': 0, 'y_offset': 0})
        draw_x = furigana_base_x + adj.get('x_offset', 0)
        draw_y = current_y + adj.get('y_offset', 0)

        # ★太字処理
        if USE_BOLD:
            canvas.setLineWidth(furigana_font_size * BOLD_WIDTH_RATIO)
            canvas.setStrokeColorRGB(0, 0, 0)
            canvas._code.append('2 Tr')

        if char in ROTATED_CHARS:
            canvas.saveState()
            canvas.translate(draw_x, draw_y)
            canvas.rotate(270)
            canvas.drawString(0, -furigana_font_size / 4, char)
            canvas.restoreState()
        else:
            canvas.drawString(draw_x, draw_y, char)

        # ★太字リセット
        if USE_BOLD:
            canvas._code.append('0 Tr')

        current_y -= furigana_font_size * char_spacing

def draw_vertical_text_with_wrap(canvas, text, x, y, font_size=12, char_spacing=1.2, max_chars_per_line=20, line_spacing=25, adjustments=None):
    if adjustments is None:
        adjustments = PUNCTUATION_ADJUSTMENTS
    
    canvas.setFont(DEFAULT_FONT, font_size)
    
    current_x, current_y = x, y
    char_count = 0
    text_str = str(text)

    for i, char in enumerate(text_str):
        if char == '\n':
            current_x -= line_spacing
            current_y = y
            char_count = 0
            continue

        is_overflow = char_count >= max_chars_per_line
        is_hanging = is_overflow and char in LEADING_PROHIBITED_CHARS

        if is_overflow and not is_hanging:
            current_x -= line_spacing
            current_y = y
            char_count = 0
        
        if is_hanging and char in HANGING_PUNCTUATION_ADJUSTMENTS:
            adj = HANGING_PUNCTUATION_ADJUSTMENTS.get(char)
        else:
            adj = adjustments.get(char, {'x_offset': 0, 'y_offset': 0})
        
        draw_x = current_x + adj['x_offset']
        draw_y = current_y + adj['y_offset']

        if is_hanging:
            draw_y += (font_size * char_spacing)
        
        # フォント切替
        target_font = SPECIAL_CHARS_FONT_MAP.get(char, DEFAULT_FONT)
        canvas.setFont(target_font, font_size)

        # ★太字処理
        if USE_BOLD:
            canvas.setLineWidth(font_size * BOLD_WIDTH_RATIO)
            canvas.setStrokeColorRGB(0, 0, 0)
            canvas._code.append('2 Tr')
        
        rotation_angle = adj.get('angle')
        if rotation_angle is not None:
            canvas.saveState()
            canvas.translate(draw_x, draw_y)
            canvas.rotate(rotation_angle)
            canvas.drawString(0, -font_size / 4, char)
            canvas.restoreState()
        elif char in ROTATED_CHARS:
            canvas.saveState()
            canvas.translate(draw_x, draw_y)
            canvas.rotate(270)
            canvas.drawString(0, -font_size / 4, char)
            canvas.restoreState()
        else:
            canvas.drawString(draw_x, draw_y, char)

        # ★太字リセット
        if USE_BOLD:
            canvas._code.append('0 Tr')

        if is_hanging:
            current_x -= line_spacing
            current_y = y
            char_count = 0
        else:
            current_y -= font_size * char_spacing
            char_count += 1

def draw_vertical_text(canvas, text, x, y, font_size=12, char_spacing=1.2, adjustments=None):
    if adjustments is None:
        adjustments = PUNCTUATION_ADJUSTMENTS
    
    for i, char in enumerate(str(text)):
        target_font = SPECIAL_CHARS_FONT_MAP.get(char, DEFAULT_FONT)
        canvas.setFont(target_font, font_size)

        adj = adjustments.get(char, {'x_offset': 0, 'y_offset': 0})
        draw_x = x + adj['x_offset']
        draw_y = y - i * font_size * char_spacing + adj['y_offset']
        
        # ★太字処理
        if USE_BOLD:
            canvas.setLineWidth(font_size * BOLD_WIDTH_RATIO)
            canvas.setStrokeColorRGB(0, 0, 0)
            canvas._code.append('2 Tr')
        
        rotation_angle = adj.get('angle')
        if rotation_angle is not None:
            canvas.saveState()
            canvas.translate(draw_x, draw_y)
            canvas.rotate(rotation_angle)
            canvas.drawString(0, -font_size / 4, char)
            canvas.restoreState()
        elif char in ROTATED_CHARS:
            canvas.saveState()
            canvas.translate(draw_x, draw_y)
            canvas.rotate(270)
            canvas.drawString(0, -font_size / 4, char)
            canvas.restoreState()
        else:
            canvas.drawString(draw_x, draw_y, char)
            
        # ★太字リセット
        if USE_BOLD:
            canvas._code.append('0 Tr')

def draw_horizontal_text_with_wrap(canvas, text, x, y, font_name, font_size=12, max_chars_per_line=40, line_height=14, adjustments=None):
    canvas.setFont(font_name, font_size)
    
    # ★太字処理
    if USE_BOLD:
        canvas.setLineWidth(font_size * BOLD_WIDTH_RATIO)
        canvas.setStrokeColorRGB(0, 0, 0)
        canvas._code.append('2 Tr')
        
    lines = []
    text_str = str(text)
    
    start = 0
    while start < len(text_str):
        end = start + max_chars_per_line
        if end >= len(text_str):
            line = text_str[start:]
            start = len(text_str)
        else:
            newline_pos = text_str.find('\n', start, end)
            if newline_pos != -1:
                line = text_str[start:newline_pos]
                start = newline_pos + 1
            else:
                if text_str[end] in LEADING_PROHIBITED_CHARS:
                    end += 1
                line = text_str[start:end]
                start = end
        
        lines.append(line)

    for i, line in enumerate(lines):
        canvas.drawString(x, y - i*line_height, line)

    # ★太字リセット
    if USE_BOLD:
        canvas._code.append('0 Tr')

def draw_centered_horizontal_text_with_wrap(canvas, text, x, y, font_name, font_size=12, max_chars_per_line=40, line_height=14, adjustments=None):
    canvas.setFont(font_name, font_size)
    
    # ★太字処理
    if USE_BOLD:
        canvas.setLineWidth(font_size * BOLD_WIDTH_RATIO)
        canvas.setStrokeColorRGB(0, 0, 0)
        canvas._code.append('2 Tr')
        
    lines = []
    text_str = str(text)

    start = 0
    while start < len(text_str):
        end = start + max_chars_per_line
        if end >= len(text_str):
            line = text_str[start:]
            start = len(text_str)
        else:
            newline_pos = text_str.find('\n', start, end)
            if newline_pos != -1:
                line = text_str[start:newline_pos]
                start = newline_pos + 1
            else:
                if text_str[end] in LEADING_PROHIBITED_CHARS:
                    end += 1
                line = text_str[start:end]
                start = end
        
        lines.append(line)

    for i, line in enumerate(lines):
        line_width = pdfmetrics.stringWidth(line, font_name, font_size)
        draw_x = x - line_width / 2
        draw_y = y - i * line_height
        canvas.drawString(draw_x, draw_y, line)
    # ★太字リセット
    if USE_BOLD:
        canvas._code.append('0 Tr')

def draw_horizontal_text(canvas, text, x, y, font_name, font_size=12):
    canvas.setFont(font_name, font_size)
    
    # ★太字処理
    if USE_BOLD:
        canvas.setLineWidth(font_size * BOLD_WIDTH_RATIO)
        canvas.setStrokeColorRGB(0, 0, 0)
        canvas._code.append('2 Tr')
        
    canvas.drawString(x, y, str(text))
    
    # ★太字リセット
    if USE_BOLD:
        canvas._code.append('0 Tr')

# ==========================================
# データ処理関数
# ==========================================

def to_full_width(text):
    if not isinstance(text, str):
        return text
    
    # 1. ASCII conversion (0x21-0x7E -> 0xFF01-0xFF5E) and Space (0x20 -> 0x3000)
    trans_table = {0x20: 0x3000}
    for i in range(0x21, 0x7F):
        trans_table[i] = i + 0xFEE0
    
    text = text.translate(trans_table)
    
    # 2. Half-width Katakana conversion using Regex and NFKC
    def replace_katakana(match):
        return unicodedata.normalize('NFKC', match.group(0))
    
    # Range for Half-width Katakana: \uff61-\uff9f
    text = re.sub(r'[\uff61-\uff9f]+', replace_katakana, text)
            
    return text

def preprocess_data(data: pd.DataFrame) -> pd.DataFrame:
    # Apply full-width conversion ONLY to '臨書解説' column if it exists
    if '臨書解説' in data.columns:
        data['臨書解説'] = data['臨書解説'].fillna('').astype(str).apply(to_full_width)

    data = data.fillna('')

    def format_df(data):
        for idx, row in data.iterrows():
            work_type = str(row.get("作品形式", "") or "")
            if work_type == "臨":
                data.at[idx, '釈文'] = row['釈文（臨書）']
                data.at[idx, '作品名'] = row['作品名（臨書）']
                data.at[idx, 'コメント'] = row['コメント（臨書）']
            else:
                data.at[idx, '釈文'] = row['釈文（創作）']
                data.at[idx, '作品名'] = row['作品名（創作）']
                data.at[idx, 'コメント'] = row['コメント（創作）']
        return data
    
    data = format_df(data)


    def combine_work_name_for_exp(row):
        work_type = str(row.get("作品形式", "") or "")
        work_name = str(row.get("作品名", "") or "")
        writer_name = str(row.get("作者名", "") or "")
        
        if writer_name == "無し":
            return f"「{work_name}」"
    
        if work_type == "臨":
            return f"{writer_name}「{work_name}」"
        return ""
    
    data["作品情報（法帖解説）"] = data.apply(combine_work_name_for_exp, axis=1)
    data.insert(0, "作品情報（法帖解説）", data.pop("作品情報（法帖解説）"))

    if "タイムスタンプ" in data.columns:
        data.drop(columns=["タイムスタンプ"], inplace=True)

    if "学部" in data.columns and "学年" in data.columns:
        data["学部学年"] = data.apply(lambda r: f"{str(r.get('学部', '') or '')} {str(r.get('学年', '') or '')}", axis=1)
        data.drop(columns=["学部", "学年"], inplace=True)
        data.insert(2, "学部学年", data.pop("学部学年"))

    if "作品名" in data.columns and "作品形式" in data.columns:
        def combine_work_name(row):
            work_name = str(row.get("作品名", "") or "")
            work_type = str(row.get("作品形式", "") or "")
            work_writer = str(row.get("作者名", "") or "")
            creation_type  = str(row.get("創作の種類", "") or "")
            if work_writer == "":
                return f"{creation_type}　「{work_name}」"
            elif work_writer == "無し":
                return f"{work_type}　「{work_name}」"
            else:
                return f"{work_type}　{work_writer}　「{work_name}」"
        data["作品情報"] = data.apply(combine_work_name, axis=1)
        drop_cols = [col for col in ["作品形式", "作品名", "作者名"] if col in data.columns]
        data.drop(columns=drop_cols, inplace=True)
        data.insert(3, "作品情報", data.pop("作品情報"))
    
    def hurigana(row):
            hurigana = str(row.get("ふりがな", "") or "")
            return f"（{hurigana}）"
    data["ふりがな"] = data.apply(hurigana, axis=1)
    data.insert(5, "ふりがな", data.pop("ふりがな"))

    def resubmission(row):
        resub = str(row.get("再提出", "") or "")
        if resub == "再提出":
            return "再提出"
        elif resub == "２回以上":
            return "やばいよ"
        else:
            return ""
    data["再提出"] = data.apply(resubmission, axis=1)
    data.insert(4, "再提出", data.pop("再提出"))

    return data


def calculate_wrap_count(text, max_chars):
    if not text:
        return 0
    
    lines = 1
    char_count = 0
    text_str = str(text)
    
    for i, char in enumerate(text_str):
        if char == '\n':
            lines += 1
            char_count = 0
            continue
        
        is_overflow = char_count >= max_chars
        is_hanging = is_overflow and char in LEADING_PROHIBITED_CHARS
        
        if is_overflow and not is_hanging:
            lines += 1
            char_count = 0
        
        if is_hanging:
            if i < len(text_str) - 1:
                lines += 1
                char_count = 0
        else:
            char_count += 1
            
    return lines - 1


def draw_content_blocks(page_canvas, data_row, x_offset=0):
    # ★枠線を描画する前に、線の設定をリセット（太さ1、黒色）して保存
    page_canvas.saveState()
    page_canvas.setLineWidth(1)
    page_canvas.setStrokeColorRGB(0, 0, 0)
    
    # --- 枠線の描画 ---
    page_canvas.rect(20 + x_offset, 20, 260, 190)
    page_canvas.rect(20 + x_offset, 230, 260, 580)
    page_canvas.line(120 + x_offset, 230, 120 + x_offset, 810)
    page_canvas.line(192 + x_offset, 230, 192 + x_offset, 810)
    page_canvas.line(236 + x_offset, 230, 236 + x_offset, 810)
    page_canvas.line(236 + x_offset, 520, 280 + x_offset, 520)
    
    # ★枠線を描き終わったら設定を元に戻す
    page_canvas.restoreState()

    for column, coord in COORDINATES.items():
        if coord.get('handler') == 'name_and_furigana':
            name_text = data_row.get("氏名", "")
            furigana_text = data_row.get("ふりがな", "")
            draw_name_and_furigana(
                page_canvas,
                name=name_text,
                furigana=furigana_text,
                x=coord['x'] + x_offset,
                y=coord['y'],
                name_font_size=coord['name_font_size'],
                furigana_font_size=coord['furigana_font_size'],
                char_spacing=coord['char_spacing'],
                adjustments=coord['adjustments'],
                furigana_x_offset=coord.get('furigana_x_offset', 0),
                furigana_y_spacing=coord.get('furigana_y_spacing', coord['name_font_size'])
            )
            continue

        value = data_row.get(column, "")
        y = coord['y']
        
        adjustments = coord.get('adjustments', PUNCTUATION_ADJUSTMENTS)
        
        if column in ['コメント', '釈文'] and isinstance(coord['x'], list):
            max_chars = coord.get('max_chars', 1)
            wrap_count = calculate_wrap_count(value, max_chars)

            x_options = coord['x']
            if wrap_count < len(x_options):
                base_x = x_options[wrap_count]
            else:
                base_x = x_options[-1]
            x = base_x + x_offset
        else:
            x = coord['x'] + x_offset

        if coord.get('horizontal', False):
            if coord.get('centered', False):
                draw_centered_horizontal_text_with_wrap(
                    page_canvas, value, x, y, DEFAULT_FONT,
                    coord['font_size'], coord.get('max_chars', 40),
                    coord.get('line_height', int(coord['font_size'] * 1.2)),
                    adjustments=adjustments
                )
            elif coord.get('wrap', False):
                draw_horizontal_text_with_wrap(
                    page_canvas, value, x, y, DEFAULT_FONT,
                    coord['font_size'], coord.get('max_chars', 40),
                    coord.get('line_height', int(coord['font_size'] * 1.2)),
                    adjustments=adjustments
                )
            else:
                draw_horizontal_text(page_canvas, value, x, y, DEFAULT_FONT, coord['font_size'])
        else:
            if coord.get('wrap', False):
                draw_vertical_text_with_wrap(
                    page_canvas, value, x, y,
                    coord['font_size'], coord['char_spacing'],
                    coord.get('max_chars', 55), coord.get('line_spacing', 25),
                    adjustments=adjustments
                )
            else:
                draw_vertical_text(page_canvas, value, x, y,
                                       coord['font_size'], coord['char_spacing'],
                                       adjustments=adjustments)

def generate_combined_pdf(data, pdf_file_path):
    page_canvas = canvas.Canvas(pdf_file_path, pagesize=portrait(A4))
    i = 0
    while i < len(data):
        if i > 0:
            page_canvas.showPage()
        
        draw_content_blocks(page_canvas, data.iloc[i], x_offset=0)
        
        if i + 1 < len(data):
            draw_content_blocks(page_canvas, data.iloc[i + 1], x_offset=OFFSET_X)
        
        i += 2
    page_canvas.save()
    messagebox.showinfo("完了", f"PDFファイルを作成しました:\n{pdf_file_path}")

def generate_individual_pdfs(data, output_dir):
    for index, row in data.iterrows():
        student_name = row.get("氏名", f"no_name_{index}")
        work_name= row.get("作品情報", f"no_work_{index}")
        safe_filename = re.sub(r'[\\/*?:"<>|]', "", student_name+" "+work_name) + ".pdf"
        pdf_path = os.path.join(output_dir, safe_filename)

        page_canvas = canvas.Canvas(pdf_path, pagesize=portrait(A4))
        draw_content_blocks(page_canvas, row, x_offset=0)
        page_canvas.save()

    messagebox.showinfo("完了", f"{len(data)}件のPDFファイルを作成しました:\n{output_dir}")

def main():
    root = tk.Tk()
    root.withdraw()

    csv_path = filedialog.askopenfilename(
        title="CSVファイルを選択してください",
        filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
    )
    if not csv_path:
        return

    try:
        data = pd.read_csv(csv_path, encoding='utf-8')
    except FileNotFoundError:
        messagebox.showerror("エラー", "CSVファイルが見つかりません。")
        return
    except UnicodeDecodeError:
        data = pd.read_csv(csv_path, encoding='shift-jis')
    except Exception as e:
        messagebox.showerror("エラー", f"CSV読み込みでエラー:\n{e}")
        return

    data = preprocess_data(data)
    
    try:
        pdfmetrics.registerFont(TTFont("HGRGE", "C:/Windows/Fonts/HGRGE.TTC"))
        pdfmetrics.registerFont(TTFont("HGRME", "C:/Windows/Fonts/HGRME.TTC"))
        pdfmetrics.registerFont(TTFont("MSMincho", "C:/Windows/Fonts/msmincho.ttc")) 
        
    except Exception as e:
        messagebox.showerror("フォントエラー", f"フォントの読み込みに失敗しました。\nC:/Windows/Fonts/ に HGRGE.TTC と HGRME.TTC、msmincho.ttc が存在するか確認してください。\n\nエラー詳細:\n{e}")
        return

    create_individual = messagebox.askyesno(
        "生成モードの選択",
        "はい(Yes): 1人ずつ個別のPDFを作成します。\nいいえ(No): 全員を1つのPDFにまとめて作成します。"
    )

    if create_individual:
        csv_dir = os.path.dirname(csv_path)
        output_dir = os.path.join(csv_dir, "nameplate")
        os.makedirs(output_dir, exist_ok=True)
        generate_individual_pdfs(data, output_dir)
    else:
        pdf_path = filedialog.asksaveasfilename(
            title="PDF保存先を選択してください",
            defaultextension=".pdf",
            filetypes=[("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")]
        )
        if pdf_path:
            generate_combined_pdf(data, pdf_path)

if __name__ == "__main__":
    main()
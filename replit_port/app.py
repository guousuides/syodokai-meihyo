from flask import Flask, render_template, request, send_file
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, portrait
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
import os
import re
import unicodedata
import io

app = Flask(__name__)

# --- Constants & Settings (Copied from applied.py) ---
USE_BOLD = True
BOLD_WIDTH_RATIO = 0.05
SPECIAL_CHARS_FONT_MAP = {'嵗': 'MSMincho', '俻': 'MSMincho'}
DEFAULT_FONT = "MSMincho"
LEADING_PROHIBITED_CHARS = {'。', '、', '」', '』', ')', '）', ']', '}'}

# Adjustments (Simplified for brevity, ideally copy full dicts)
HANGING_PUNCTUATION_ADJUSTMENTS = {'。': {'x_offset': 8, 'y_offset': -2}, '、': {'x_offset': 8, 'y_offset': -2}}
PUNCTUATION_ADJUSTMENTS = {} # Populate with full data if needed for perfect fidelity
ROTATED_CHARS = {'(', ')', '（', '）', '[', ']', '{', '}', '「', '」', '『', '』', 'ー','-','→','←','↑','↓'}

# Coordinates (Copied from applied.py)
COORDINATES = {
    '学部学年': {'x': 250, 'y': 780, 'font_size': 18, 'char_spacing': 1.0, 'wrap': False, 'horizontal': False},
    '名前': {'handler': 'name_and_furigana', 'x': 250, 'y': 495, 'name_font_size': 18, 'furigana_font_size': 13, 'furigana_x_offset': 2.75, 'furigana_y_spacing': 0.5, 'char_spacing': 1.0},
    '作品情報': {'x': 206, 'y': 780, 'font_size': 18, 'char_spacing': 1.0, 'wrap': False, 'horizontal': False},
    '釈文': {'x': [152, 162, 172], 'y': 786, 'font_size': 10, 'char_spacing': 1.0, 'wrap': True, 'max_chars': 55, 'line_spacing': 20, 'horizontal': False},
    'コメント': {'x': [70, 80, 85, 95, 105], 'y': 786, 'font_size': 10, 'char_spacing': 1.0, 'wrap': True, 'max_chars': 55, 'line_spacing': 20, 'horizontal': False},
    '臨書解説': {'x': 25, 'y': 170, 'font_size': 10, 'char_spacing': 1, 'wrap': True, 'max_chars': 25, 'line_spacing': 30, 'horizontal': True},
    '作品情報（法帖解説）': {'x': 150, 'y': 190, 'font_size': 14, 'char_spacing': 1, 'wrap': True, 'max_chars': 25, 'line_spacing': 30, 'horizontal': True, 'centered': True},
    '再提出': {'x': 150, 'y': 5, 'font_size': 14, 'char_spacing': 1, 'wrap': True, 'max_chars': 25, 'line_spacing': 30, 'horizontal': True, 'centered': True},
}
OFFSET_X = 300

# --- Helper Functions (Simplified versions) ---
def to_full_width(text):
    if not isinstance(text, str): return text
    trans_table = {0x20: 0x3000}
    for i in range(0x21, 0x7F): trans_table[i] = i + 0xFEE0
    return text.translate(trans_table)

def preprocess_data(data: pd.DataFrame) -> pd.DataFrame:
    # (Copy logic from applied.py)
    if '臨書解説' in data.columns:
        data['臨書解説'] = data['臨書解説'].fillna('').astype(str).apply(to_full_width)
    data = data.fillna('')
    
    # ... (Include other preprocessing steps as needed) ...
    # For simulation, we assume input is already somewhat structured, but let's keep the logic
    def format_df(data):
        for idx, row in data.iterrows():
            work_type = str(row.get("作品形式", "") or "")
            if work_type == "臨":
                data.at[idx, '釈文'] = row.get('釈文（臨書）', '')
                data.at[idx, '作品名'] = row.get('作品名（臨書）', '')
                data.at[idx, 'コメント'] = row.get('コメント（臨書）', '')
            else:
                data.at[idx, '釈文'] = row.get('釈文（創作）', '')
                data.at[idx, '作品名'] = row.get('作品名（創作）', '')
                data.at[idx, 'コメント'] = row.get('コメント（創作）', '')
        return data
    data = format_df(data)
    
    # Combine fields logic (simplified for brevity)
    if "学部" in data.columns and "学年" in data.columns:
        data["学部学年"] = data.apply(lambda r: f"{str(r.get('学部', '') or '')} {str(r.get('学年', '') or '')}", axis=1)
    
    # ... (Other combinations) ...
    
    return data

# --- Drawing Functions (Placeholder for full implementation) ---
# In a real port, you would copy the full draw_name_and_furigana, draw_vertical_text_with_wrap etc.
# Here I will use simplified versions to demonstrate the structure.

def draw_content_blocks(page_canvas, data_row, x_offset=0):
    # Draw lines
    page_canvas.saveState()
    page_canvas.setLineWidth(1)
    page_canvas.rect(20 + x_offset, 20, 260, 190)
    page_canvas.rect(20 + x_offset, 230, 260, 580)
    page_canvas.line(120 + x_offset, 230, 120 + x_offset, 810)
    page_canvas.line(192 + x_offset, 230, 192 + x_offset, 810)
    page_canvas.line(236 + x_offset, 230, 236 + x_offset, 810)
    page_canvas.line(236 + x_offset, 520, 280 + x_offset, 520)
    page_canvas.restoreState()

    # Draw text (Simplified)
    page_canvas.setFont(DEFAULT_FONT, 12)
    for column, coord in COORDINATES.items():
        value = data_row.get(column, "")
        if not value: continue
        
        # Just drawing text at x,y for demo. 
        # You MUST copy the full logic from applied.py for vertical text support.
        page_canvas.drawString(coord['x'] + x_offset, coord['y'], str(value))


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.form.to_dict()
    
    # Map form fields to DataFrame expected by preprocess_data
    # Note: The form should send keys matching what preprocess_data expects
    # Or we manually map them here.
    
    # For simplicity, let's assume the form sends '氏名', 'ふりがな', etc.
    # And we handle the '作品形式' logic here or in preprocess.
    
    df = pd.DataFrame([data])
    # processed_df = preprocess_data(df) # Call full preprocess
    row = df.iloc[0] # Use processed row
    
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=portrait(A4))
    
    # Register fonts (Ensure font files are available in Replit)
    try:
        pdfmetrics.registerFont(TTFont("MSMincho", "msmincho.ttc")) # You need to upload this font
    except:
        pass # Fallback
        
    draw_content_blocks(p, row)
    p.save()
    buffer.seek(0)
    
    return send_file(buffer, as_attachment=True, download_name="simulation.pdf", mimetype='application/pdf')

@app.route('/save_csv', methods=['POST'])
def save_csv():
    data = request.form.to_dict()
    df = pd.DataFrame([data])
    
    # Ensure columns are in a logical order if possible, or just dump as is
    # For better UX, we might want to order them: 氏名, ふりがな, 学部, etc.
    preferred_order = ['氏名', 'ふりがな', '学部', '作品情報', '釈文', 'コメント']
    # Add any extra columns that might be in data but not in preferred_order
    cols = [c for c in preferred_order if c in df.columns] + [c for c in df.columns if c not in preferred_order]
    df = df[cols]
    
    buffer = io.BytesIO()
    # BOM for Excel compatibility in Japanese environment
    buffer.write(b'\xef\xbb\xbf')
    df.to_csv(buffer, index=False, encoding='utf-8')
    buffer.seek(0)
    
    return send_file(buffer, as_attachment=True, download_name="input_data.csv", mimetype='text/csv')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)

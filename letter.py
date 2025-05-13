from flask import Flask, request, send_file, render_template_string
import os
import fitz  # PyMuPDF
import zipfile
from io import BytesIO

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# หัวกระดาษใหม่ที่ต้องแทนที่
new_header_lines = [
    "บริษัท เวิร์ค สเตชั่น ออฟฟิศ (ประเทศไทย) จำกัด",
    "ที่อยู่: 440 442 พระราม 2 ซ. 50 แขวงแสมดำ เขตบางขุนเทียน กรุงเทพมหานคร 10150",
    "โทรศัพท์: 021004740, อีเมล: inventory@workstationoffice.com",
    "เลขผู้เสียภาษี: 0105552123122"
]

header_settings = {
    "tax":     {"rect": (0, 0, 425, 100), "text_x": 30, "text_y": 40, "fontsize": 11, "line_spacing": 18},
    "receipt": {"rect": (0, 0, 425, 100), "text_x": 30, "text_y": 40, "fontsize": 11, "line_spacing": 18},
    "both":    {"rect": (0, 0, 425, 100), "text_x": 30, "text_y": 40, "fontsize": 11, "line_spacing": 18},
}

# HTML Template แบบปรับหน้าตาให้ดูดีขึ้นด้วย CSS
HTML_TEMPLATE = '''
<!doctype html>
<html lang="th">
<head>
    <meta charset="utf-8">
    <title>เปลี่ยนหัวกระดาษ PDF</title>
    <style>
        body {
            font-family: sans-serif;
            background-color: #f4f4f4;
            padding: 30px;
            color: #333;
        }
        h1 {
            text-align: center;
            color: #2c3e50;
        }
        .section {
            background: white;
            border-radius: 8px;
            padding: 20px;
            margin: 20px auto;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            max-width: 600px;
        }
        input[type=file] {
            width: 100%;
            padding: 8px;
            margin-bottom: 10px;
        }
        input[type=submit] {
            background-color: #3498db;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
        }
        input[type=submit]:hover {
            background-color: #2980b9;
        }
    </style>
</head>
<body>
    <h1>เครื่องมือเปลี่ยนหัวกระดาษ PDF</h1>

    <div class="section">
        <h2>สำหรับ ใบกํากับภาษี</h2>
        <form action="/tax" method=post enctype=multipart/form-data>
            <input type=file name=pdfs multiple required>
            <input type=submit value="แปลงไฟล์">
        </form>
    </div>

    <div class="section">
        <h2>สำหรับ ใบเสร็จรับเงิน</h2>
        <form action="/receipt" method=post enctype=multipart/form-data>
            <input type=file name=pdfs multiple required>
            <input type=submit value="แปลงไฟล์">
        </form>
    </div>

    <div class="section">
        <h2>สำหรับ ใบกํากับภาษี/ใบเสร็จรับเงิน</h2>
        <form action="/both" method=post enctype=multipart/form-data>
            <input type=file name=pdfs multiple required>
            <input type=submit value="แปลงไฟล์">
        </form>
    </div>
</body>
</html>
'''

@app.route('/', methods=['GET'])
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/tax', methods=['POST'])
def handle_tax():
    return handle_conversion("tax")

@app.route('/receipt', methods=['POST'])
def handle_receipt():
    return handle_conversion("receipt")

@app.route('/both', methods=['POST'])
def handle_both():
    return handle_conversion("both")

def handle_conversion(mode):
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        for uploaded_file in request.files.getlist("pdfs"):
            if uploaded_file.filename.endswith('.pdf'):
                filepath = os.path.join(UPLOAD_FOLDER, uploaded_file.filename)
                uploaded_file.save(filepath)

                new_pdf_path = process_pdf(filepath, mode)
                arcname = f"WS_{os.path.basename(new_pdf_path)}"
                zipf.write(new_pdf_path, arcname)

    zip_buffer.seek(0)
    return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name=f'converted_{mode}.zip')

def process_pdf(pdf_path, mode):
    doc = fitz.open(pdf_path)
    first_page = doc[0]

    config = header_settings.get(mode, header_settings["both"])
    rect = fitz.Rect(*config["rect"])
    first_page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))

    font_path = "Sarabun-Thin.ttf"
    first_page.insert_font("THSarabun", fontfile=font_path)

    y = config["text_y"]
    for line in new_header_lines:
        first_page.insert_text((config["text_x"], y), line, fontsize=config["fontsize"], fontname="THSarabun")
        y += config["line_spacing"]

    output_path = os.path.join(OUTPUT_FOLDER, os.path.basename(pdf_path))
    doc.save(output_path)
    doc.close()
    return output_path

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=8080, debug=True)

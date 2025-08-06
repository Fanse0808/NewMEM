import os
import shutil
import zipfile
import tempfile
from flask import Flask, render_template, request, redirect, flash, send_file
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'supersecretkey')

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
FONT_PATH = 'static/Montserrat-Bold.ttf'  # ✅ adjusted for your folder

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# 🔁 Cache font once
try:
    default_font = ImageFont.truetype(FONT_PATH, 40)
except:
    default_font = ImageFont.load_default()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        flash('No file part')
        return redirect('/')
    
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect('/')

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    try:
        # 🔽 faster Excel/CSV parsing
        if filepath.endswith('.csv'):
            df = pd.read_csv(filepath, dtype=str)
        elif filepath.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(filepath, engine='openpyxl', dtype=str)
        else:
            flash('Unsupported file type.')
            return redirect('/')

        if df.empty:
            flash('Uploaded file is empty.')
            return redirect('/')

        # 💡 safer clean-up
        if os.path.exists(OUTPUT_FOLDER):
            for f in os.listdir(OUTPUT_FOLDER):
                os.remove(os.path.join(OUTPUT_FOLDER, f))

        generate_cards_from_df(df, OUTPUT_FOLDER)

        # 🔧 Create ZIP in memory-friendly temp dir
        zip_path = os.path.join(tempfile.gettempdir(), 'cards.zip')
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for filename in os.listdir(OUTPUT_FOLDER):
                file_path = os.path.join(OUTPUT_FOLDER, filename)
                zipf.write(file_path, arcname=filename)

        return send_file(zip_path, as_attachment=True)

    except Exception as e:
        flash(f'Error processing file: {str(e)}')
        return redirect('/')

def generate_cards_from_df(df, output_folder):
    for i, row in df.iterrows():
        img = Image.new('RGB', (1000, 600), color='white')
        draw = ImageDraw.Draw(img)

        # 🧠 Use cached font
        font = default_font

        # ✏️ Extract fields with defaults
        name = str(row.get('Name', '')).strip()
        email = str(row.get('Email', '')).strip()
        title = str(row.get('Title', row.get('Department', ''))).strip()
        card_number = str(row.get('Card Number', row.get('ID', f"Card{i+1}"))).strip()

        # 🖼️ Draw text
        draw.text((50, 80), f"Card #: {card_number}", fill='black', font=font)
        draw.text((50, 160), f"Name: {name}", fill='black', font=font)
        draw.text((50, 240), f"Email: {email}", fill='black', font=font)
        draw.text((50, 320), f"Title: {title}", fill='black', font=font)

        # 💾 Save image
        filename = f"{card_number.replace('/', '_')}.png"
        img.save(os.path.join(output_folder, filename))

if __name__ == '__main__':
    app.run(debug=True)

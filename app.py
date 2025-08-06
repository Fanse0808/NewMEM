import os
import shutil
import zipfile
import tempfile
import logging
import traceback
from flask import Flask, render_template, request, redirect, flash, send_file, after_this_request
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Setup
app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'your_secret_key')

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'static/output'
FONT_PATH = 'static/fonts/Montserrat-Bold.ttf'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

logging.basicConfig(level=logging.INFO)

ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

REQUIRED_COLUMNS = ['Name', 'Email', 'Title', 'Card Number']


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def validate_columns(df):
    return all(col in df.columns for col in REQUIRED_COLUMNS)


def generate_cards_from_df(df, output_folder):
    for i, row in df.iterrows():
        img = Image.new('RGB', (1000, 600), color='white')
        d = ImageDraw.Draw(img)

        try:
            font = ImageFont.truetype(FONT_PATH, 40)
        except:
            font = ImageFont.load_default()

        name = str(row.get('Name', '')).strip()
        email = str(row.get('Email', '')).strip()
        title = str(row.get('Title', row.get('Department', ''))).strip()
        card_number = str(row.get('Card Number', row.get('ID', f"Card{i+1}"))).strip()

        d.text((50, 80), f"Card #: {card_number}", fill='black', font=font)
        d.text((50, 160), f"Name: {name}", fill='black', font=font)
        d.text((50, 240), f"Email: {email}", fill='black', font=font)
        d.text((50, 320), f"Title: {title}", fill='black', font=font)

        # Safe filename
        safe_filename = f"{card_number.replace('/', '_') or f'Card_{i+1}'}.png"
        img.save(os.path.join(output_folder, safe_filename))


def zip_folder(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                zipf.write(os.path.join(root, file),
                           arcname=os.path.relpath(os.path.join(root, file), folder_path))


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file')

        if not file or file.filename == '':
            flash('No file uploaded.')
            return redirect(request.url)

        if allowed_file(file.filename):
            try:
                shutil.rmtree(OUTPUT_FOLDER, ignore_errors=True)
                os.makedirs(OUTPUT_FOLDER, exist_ok=True)

                filepath = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(filepath)

                # Load CSV or Excel
                if file.filename.endswith('.xlsx'):
                    df = pd.read_excel(filepath, engine='openpyxl')
                else:
                    try:
                        df = pd.read_csv(filepath, encoding='utf-8')
                    except UnicodeDecodeError:
                        df = pd.read_csv(filepath, encoding='latin1')

                logging.info("Headers detected: %s", df.columns.tolist())

                # Check required columns
                if not validate_columns(df):
                    flash(f"Missing required columns. Required: {', '.join(REQUIRED_COLUMNS)}")
                    return redirect(request.url)

                generate_cards_from_df(df, OUTPUT_FOLDER)

                with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as tmp_zip:
                    zip_folder(OUTPUT_FOLDER, tmp_zip.name)
                    zip_path = tmp_zip.name

                @after_this_request
                def cleanup(response):
                    try:
                        os.remove(zip_path)
                        logging.info("Temp zip deleted")
                    except Exception as e:
                        logging.error("Error deleting zip: %s", e)
                    return response

                return send_file(zip_path, as_attachment=True, download_name='cards.zip')

            except Exception as e:
                logging.error("Error: %s", traceback.format_exc())
                flash(f"An error occurred while processing the file: {str(e)}")
                return redirect(request.url)

        flash('Unsupported file type. Use .csv or .xlsx only.')
        return redirect(request.url)

    return render_template('index.html')

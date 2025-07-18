

import os
import shutil
import zipfile
import tempfile
import re
import threading
import time
import logging
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Logging for Render troubleshooting
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'supersecretkey')
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'web_cards'
TEMPLATE_IMAGE = 'Card_Regular.jpg'
FONT_PATH = 'arial.ttf'
LEFT_MARGIN = 60
POLICY_NO_POS = (LEFT_MARGIN, 230)
VALID_UNTIL_LABEL_POS = (LEFT_MARGIN, 280)
NAME_POS = (LEFT_MARGIN, 400)
FONT_SIZE_POLICY_NO = 36.4
FONT_SIZE_DATE = 22
FONT_SIZE_NAME = 25
FONT_SIZE_LABEL = 18
WHITE = (255, 255, 255)
for folder in (UPLOAD_FOLDER, OUTPUT_FOLDER, 'static'):
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception as e:
        logging.error(f"Failed to create folder {folder}: {e}")



# Create sample CSV file if it doesn't exist (now with VIP and Email columns)
SAMPLE_CSV_PATH = 'static/sample_cards.csv'
if not os.path.exists(SAMPLE_CSV_PATH):
    sample_data = "Name,Card,Date,VIP,Email\nJohn Doe,STE 12345 690 7890,2024-12-31,Yes,john@example.com\nJane Smith,CII 98765 432 1098,2025-01-15,No,jane@example.com"
    with open(SAMPLE_CSV_PATH, 'w') as f:
        f.write(sample_data)


# Helper function to send email with attachment (using smtplib)
import smtplib
from email.message import EmailMessage
import mimetypes

def send_email_with_attachment(to_email, subject, body, attachment_path):
    smtp_server = os.environ.get('SMTP_SERVER')
    smtp_port = int(os.environ.get('SMTP_PORT', 587))
    smtp_user = os.environ.get('SMTP_USER')
    smtp_password = os.environ.get('SMTP_PASSWORD')
    from_email = smtp_user

    if not all([smtp_server, smtp_port, smtp_user]):
        logging.error("SMTP configuration is missing. Set SMTP_SERVER, SMTP_PORT, SMTP_USER as environment variables.")
        return

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = to_email
    msg.set_content(body)

    # Attach the file efficiently (stream, avoid loading whole file in memory)
    if attachment_path and os.path.exists(attachment_path):
        try:
            maintype, subtype = mimetypes.guess_type(attachment_path)[0].split('/') if mimetypes.guess_type(attachment_path)[0] else ('application', 'octet-stream')
            filename = os.path.basename(attachment_path)
            with open(attachment_path, 'rb') as f:
                msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=filename)
        except Exception as e:
            logging.error(f"Failed to attach file {attachment_path}: {e}")

    try:
        with smtplib.SMTP(smtp_server, smtp_port, timeout=10) as server:
            if smtp_password:
                server.starttls()
                server.login(smtp_user, smtp_password)
            server.send_message(msg)
        logging.info(f"Email sent to {to_email} with attachment {attachment_path}")
    except Exception as e:
        logging.error(f"Failed to send email to {to_email}: {e}")


# Place the /api/create_card endpoint after app is defined
@app.route('/api/create_card', methods=['POST'])
def api_create_card():
    data = request.get_json()
    if not data:
        return {"error": "No JSON payload provided"}, 400
    name = data.get('Name')
    Card = data.get('Card')
    date = data.get('Date')
    if not all([name, Card, date]):
        return {"error": "Missing required fields: Name, Card, Date"}, 400
    # Create a DataFrame with one row
    df = pd.DataFrame([[name, Card, date]], columns=['Name', 'Card', 'Date'])
    output_folder = tempfile.mkdtemp()
    generate_cards_from_df(df, output_folder)
    card_file = os.path.join(output_folder, f"{sanitize_filename(name)}_{sanitize_filename(Card)}.png")
    if not os.path.exists(card_file):
        # fallback if file not found
        card_file = next((os.path.join(output_folder, f) for f in os.listdir(output_folder) if f.endswith('.png')), None)
        if not card_file:
            return {"error": "Card image not generated"}, 500
    return send_file(card_file, as_attachment=True, download_name='card.png')

# --- Utility Functions ---
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'csv'}

def format_Card(Card):
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', '', str(Card))
    cleaned = ' '.join(cleaned.split())
    return cleaned

def load_font(path, size):
    try:
        return ImageFont.truetype(path, int(size))
    except IOError:
        return ImageFont.load_default()

def sanitize_filename(name):
    name = name.replace(' ', '_')
    name = re.sub(r'[\\/*?:"<>|\n]', '', name)
    return name[:100]

def generate_cards_from_df(df, output_folder):
    font_label = load_font(FONT_PATH, int(FONT_SIZE_LABEL))
    font_policy_no = load_font(FONT_PATH, int(FONT_SIZE_POLICY_NO))
    font_date = load_font(FONT_PATH, int(FONT_SIZE_DATE))
    font_name = load_font(FONT_PATH, int(FONT_SIZE_NAME))

    name_col = df.columns[0]
    Card_col = df.columns[1]
    date_col = df.columns[2]
    vip_col = df.columns[3] if len(df.columns) > 3 else None
    email_col = df.columns[4] if len(df.columns) > 4 else None

    for i, row in enumerate(df.itertuples(index=False)):
        name = str(getattr(row, name_col)) if hasattr(row, name_col) and not pd.isna(getattr(row, name_col)) else "Unknown"
        Card = str(getattr(row, Card_col)) if hasattr(row, Card_col) and not pd.isna(getattr(row, Card_col)) else "Unknown"
        date = str(getattr(row, date_col)) if hasattr(row, date_col) and not pd.isna(getattr(row, date_col)) else ""
        vip_status = str(getattr(row, vip_col)).strip().lower() if vip_col and hasattr(row, vip_col) and not pd.isna(getattr(row, vip_col)) else "no"
        email = str(getattr(row, email_col)).strip() if email_col and hasattr(row, email_col) and not pd.isna(getattr(row, email_col)) else None

        template_img = os.path.join('static', 'Card_VIP.jpg') if vip_status == 'yes' else os.path.join('static', 'Card_Regular.jpg')
        if not os.path.exists(template_img):
            template_img = TEMPLATE_IMAGE

        with Image.open(template_img) as im:
            TEMPLATE_SIZE = im.size
            card = im.convert('RGB').resize(TEMPLATE_SIZE)
            draw = ImageDraw.Draw(card)

            formatted_Card = format_Card(Card)
            draw.text(POLICY_NO_POS, formatted_Card, font=font_policy_no, fill=WHITE)

            valid_until_label = "Valid Until"
            draw.text(VALID_UNTIL_LABEL_POS, valid_until_label, font=font_label, fill=WHITE)
            bbox = draw.textbbox(VALID_UNTIL_LABEL_POS, valid_until_label, font=font_label)
            label_width = bbox[2] - bbox[0]
            date_x = VALID_UNTIL_LABEL_POS[0] + label_width + 10
            draw.text((date_x, VALID_UNTIL_LABEL_POS[1]), date, font=font_date, fill=WHITE)

            draw.text(NAME_POS, name, font=font_name, fill=WHITE)

            safe_name = sanitize_filename(name)
            safe_Card = sanitize_filename(Card)
            filename = os.path.join(output_folder, f"{safe_name}_{safe_Card}.png")

            try:
                card.save(filename, format='PNG')
            except Exception as e:
                logging.error(f"Error saving {filename}: {str(e)}")
                fallback_filename = os.path.join(output_folder, f"card_{i}.png")
                card.save(fallback_filename, format='PNG')

            if email:
                subject = f"Your A-Member Card Awaits You ({'VIP' if vip_status == 'yes' else 'Regular'})"
                body = f"Dear {name},\n\nPlease find your A-Member card attached.\n\nCustomer Care & Complaints Management\nOperation Department\n\nPhone: +95 9791232222\n\nEmail: customercare@alife.com.mm\n\nA Life Insurance Company Limited"
                send_email_with_attachment(email, subject, body, filename)

def zip_folder(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.png'):
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, arcname=file)

def clear_folders_periodically():
    while True:
        for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    logging.error(f'Failed to delete {file_path}. Reason: {e}')
        time.sleep(60 * 60 * 12)  # 12 hours

# Start the background thread for periodic cleanup (only if not running in test mode)
if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
    cleanup_thread = threading.Thread(target=clear_folders_periodically, daemon=True)
    cleanup_thread.start()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if not file.filename:
            flash('No filename provided for upload')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            try:
                shutil.rmtree(OUTPUT_FOLDER, ignore_errors=True)
                os.makedirs(OUTPUT_FOLDER, exist_ok=True)
            except Exception:
                flash('Error preparing output folder')
                return redirect(request.url)
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)
            try:
                if file.filename.lower().endswith('.xlsx'):
                    df = pd.read_excel(filepath, engine='openpyxl')
                else:
                    df = pd.read_csv(filepath)
            except Exception:
                flash('Error reading file')
                return redirect(request.url)
            if len(df.columns) < 3:
                flash('File must have at least 3 columns')
                return redirect(request.url)
            try:
                generate_cards_from_df(df, OUTPUT_FOLDER)
            except Exception as e:
                flash(f'Error generating cards: {str(e)}')
                return redirect(request.url)
            with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as tmp_zip:
                zip_folder(OUTPUT_FOLDER, tmp_zip.name)
                tmp_zip_path = tmp_zip.name
            return send_file(tmp_zip_path, as_attachment=True, download_name='cards.zip')
        else:
            flash('PLEASE DOWNLOAD THE TEMPLATE')
            return redirect(request.url)
    return render_template('index.html')

@app.route('/download_template')
def download_template():
    return send_file(SAMPLE_CSV_PATH, as_attachment=True, download_name='sample_cards.csv')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    logging.info(f"Starting Flask app on port {port}")
    app.run(host='0.0.0.0', port=port, threaded=True)
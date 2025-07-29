import os
import shutil
import zipfile
import tempfile
import re
import threading
import time
import logging
import smtplib
import mimetypes
from email import encoders
from email.message import EmailMessage
from email.utils import make_msgid
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Logging
logging.basicConfig(level=logging.INFO)

# App Setup
app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'supersecretkey')
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'web_cards'
SAMPLE_CSV_PATH = 'static/sample_cards.csv'
FONT_PATH = 'Hopone.ttf'
TEMPLATE_IMAGE = 'Card_Regular.jpg'
WHITE = (255, 255, 255)

# Text positions
LEFT_MARGIN = 60
POLICY_NO_POS = (LEFT_MARGIN, 200)
VALID_UNTIL_LABEL_POS = (LEFT_MARGIN, 280)
NAME_POS = (LEFT_MARGIN, 390)
FONT_SIZE_POLICY_NO = 36.4
FONT_SIZE_DATE = 22
FONT_SIZE_NAME = 25
FONT_SIZE_LABEL = 18

# Ensure folders exist
for folder in (UPLOAD_FOLDER, OUTPUT_FOLDER, 'static'):
    os.makedirs(folder, exist_ok=True)

# Create sample CSV
if not os.path.exists(SAMPLE_CSV_PATH):
    with open(SAMPLE_CSV_PATH, 'w') as f:
        f.write("Name,Card,Date,VIP,Email\nJohn Doe,STE 12345 690 7890,2024-12-31,Yes,john@example.com\nJane Smith,CII 98765 432 1098,2025-01-15,No,jane@example.com")

# --- Utility Functions ---
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'csv'}

def load_font(path, size):
    try:
        return ImageFont.truetype(path, int(size))
    except IOError:
        return ImageFont.load_default()

def sanitize_filename(name):
    name = name.replace(' ', '_')
    return re.sub(r'[\\/*?:"<>|\n]', '', name)[:100]

def format_card_id(card_id):
    cleaned = re.sub(r'[^a-zA-Z0-9]', '', str(card_id))
    chars = cleaned[:3].ljust(3, 'X')
    numbers = ''.join(c for c in cleaned if c.isdigit())[:11].ljust(11, '0')
    return f"{chars}-{numbers[:4]} {numbers[4:8]} {numbers[8:11]}"

import os, smtplib, mimetypes, logging
from email.message import EmailMessage
from email.utils import make_msgid

def send_email_with_attachment(to_email, subject, body_text, attachment_path=None):
    smtp_server = os.environ.get('SMTP_SERVER')
    smtp_port = int(os.environ.get('SMTP_PORT', 587))
    smtp_user = os.environ.get('SMTP_USER')
    smtp_password = os.environ.get('SMTP_PASSWORD')

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = smtp_user
    msg['To'] = to_email

    image_cid = make_msgid(domain='amember.local')[1:-1]

    contact_info = """<div style='text-align:left;'><br>Warm Regards,<br>Customer Care & Complaints Management
    <br>Operation Department<br><br>Phone: +95 9791232222<br>Email: customercare@alife.com.mm<br><br>
    A Life Insurance Company Limited<br>3rd Floor (A), No. (108), Corner of<br>
    Kabaraye Pagoda Road and Nat Mauk Road,<br>
    Bo Cho (1) Quarter, Bahan Township, Yangon, Myanmar 12201<br></div>"""

    html_body = f"""
    <html><body>
        <img src="cid:{image_cid}" style="max-width:100%;"><p>{body_text}</p>
        {contact_info}
    </body></html>
    """

    msg.set_content(body_text)
    msg.add_alternative(html_body, subtype='html')

    try:
        with open(os.path.join('static', 'memberinfo.jpg'), 'rb') as img:
            msg.get_payload()[1].add_related(
                img.read(),
                maintype='image',
                subtype='jpeg',
                cid=f"<{image_cid}>",
                disposition='inline',           # <--- Forces inline
                filename='memberinfo.jpg'       # Helps avoid 'noname'
            )
    except Exception as e:
        logging.error(f"Embed image failed: {e}")

    redemption_path = os.path.join('static', 'Redemption.jpg')
    if os.path.exists(redemption_path):
        with open(redemption_path, 'rb') as img:
            msg.add_attachment(
                img.read(),
                maintype='image',
                subtype='jpeg',
                filename='Redemption.jpg'
            )

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, 'rb') as f:
            mime_type, _ = mimetypes.guess_type(attachment_path)
            maintype, subtype = mime_type.split('/') if mime_type else ('application', 'octet-stream')
            msg.add_attachment(
                f.read(),
                maintype=maintype,
                subtype=subtype,
                filename=os.path.basename(attachment_path)
            )

    with smtplib.SMTP(smtp_server, smtp_port, timeout=10) as server:
        server.starttls()
        if smtp_password:
            server.login(smtp_user, smtp_password)
        server.send_message(msg)
        
def generate_cards_from_df(df, output_folder):
    font_label = load_font(FONT_PATH, FONT_SIZE_LABEL)
    font_policy_no = load_font(FONT_PATH, FONT_SIZE_POLICY_NO)
    font_date = load_font(FONT_PATH, FONT_SIZE_DATE)
    font_name = load_font(FONT_PATH, FONT_SIZE_NAME)

    for i, row in df.iterrows():
        name = str(row.get('Name', 'Unknown'))
        Card = str(row.get('Card', 'Unknown'))
        date = str(row.get('Date', ''))
        vip_status = str(row.get('VIP', 'no')).strip().lower()
        email = str(row.get('Email')) if pd.notna(row.get('Email')) else None

        template_img = os.path.join('static', 'Card_VIP.jpg' if vip_status == 'yes' else 'Card_Regular.jpg')
        if not os.path.exists(template_img):
            template_img = os.path.join('static', TEMPLATE_IMAGE)

        with Image.open(template_img) as im:
            card = im.convert('RGB')
            draw = ImageDraw.Draw(card)
            draw.text(POLICY_NO_POS, format_card_id(Card), font=font_policy_no, fill=WHITE)
            draw.text(VALID_UNTIL_LABEL_POS, "VALID", font=font_label, fill=WHITE)
            bbox = draw.textbbox(VALID_UNTIL_LABEL_POS, "VALID", font=font_label)
            label_height = bbox[3] - bbox[1]
            second_line_y = VALID_UNTIL_LABEL_POS[1] + label_height + 5
            until_date_text = f"UNTIL - {date}"
            draw.text((VALID_UNTIL_LABEL_POS[0], second_line_y), until_date_text, font=font_date, fill=WHITE)
            draw.text(NAME_POS, name, font=font_name, fill=WHITE)

            filename = os.path.join(output_folder, f"{sanitize_filename(name)}_{sanitize_filename(Card)}.png")
            card.save(filename, format='PNG')

            if email:
                subject = f"Your A-Member Card Awaits You ({'VIP' if vip_status == 'yes' else 'Regular'})"
                send_email_with_attachment(email, subject, "", filename)

def zip_folder(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.png'):
                    zipf.write(os.path.join(root, file), arcname=file)

def clear_folders_periodically():
    while True:
        for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    logging.error(f"Cleanup error: {e}")
        time.sleep(43200)  # every 12 hours

if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
    threading.Thread(target=clear_folders_periodically, daemon=True).start()

# Routes
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename:
            flash('No file uploaded')
            return redirect(request.url)

        if allowed_file(file.filename):
            try:
                shutil.rmtree(OUTPUT_FOLDER, ignore_errors=True)
                os.makedirs(OUTPUT_FOLDER, exist_ok=True)
                filepath = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(filepath)
                df = pd.read_excel(filepath, engine='openpyxl') if file.filename.endswith('.xlsx') else pd.read_csv(filepath)
                if df.shape[1] < 3:
                    flash('File must have at least 3 columns')
                    return redirect(request.url)
                generate_cards_from_df(df, OUTPUT_FOLDER)
                with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as tmp_zip:
                    zip_folder(OUTPUT_FOLDER, tmp_zip.name)
                return send_file(tmp_zip.name, as_attachment=True, download_name='cards.zip')
            except Exception as e:
                flash(f"Processing error: {e}")
                return redirect(request.url)
        else:
            flash('Unsupported file type. Please use the template.')
            return redirect(request.url)
    return render_template('index.html')

@app.route('/download_template')
def download_template():
    return send_file(SAMPLE_CSV_PATH, as_attachment=True, download_name='sample_cards.csv')

@app.route('/api/create_card', methods=['POST'])
def api_create_card():
    data = request.get_json()
    if not data:
        return {"error": "No JSON payload provided"}, 400
    if not all(key in data for key in ('Name', 'Card', 'Date')):
        return {"error": "Missing fields: Name, Card, Date"}, 400

    df = pd.DataFrame([[data['Name'], data['Card'], data['Date']]], columns=['Name', 'Card', 'Date'])
    output_folder = tempfile.mkdtemp()
    generate_cards_from_df(df, output_folder)
    card_file = next((os.path.join(output_folder, f) for f in os.listdir(output_folder) if f.endswith('.png')), None)
    if not card_file:
        return {"error": "Card image not generated"}, 500
    return send_file(card_file, as_attachment=True, download_name='card.png')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    logging.info(f"Starting Flask app on port {port}")
    app.run(host='0.0.0.0', port=port, threaded=True)

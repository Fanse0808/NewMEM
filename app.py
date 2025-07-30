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
import traceback
import base64
from email.message import EmailMessage
from flask import Flask, render_template, request, redirect, flash, send_file, after_this_request
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Logging
logging.basicConfig(level=logging.INFO)

# Flask Setup
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

# Create sample CSV if not exists
if not os.path.exists(SAMPLE_CSV_PATH):
    with open(SAMPLE_CSV_PATH, 'w') as f:
        f.write("Name,Card,Date,VIP,Email\nJohn Doe,STE 12345 690 7890,2024-12-31,Yes,john@example.com\n"
                "Jane Smith,CII 98765 432 1098,2025-01-15,No,jane@example.com")

# ---- Utility Functions ----
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

def send_email_with_attachment(to_email, subject, body_text, attachment_path=None):
    smtp_server = os.environ.get('SMTP_SERVER')
    smtp_port = int(os.environ.get('SMTP_PORT', 587))
    smtp_user = os.environ.get('SMTP_USER')
    smtp_password = os.environ.get('SMTP_PASSWORD')

    if not all([smtp_server, smtp_user, smtp_password]):
        logging.error("Missing SMTP environment variables.")
        return

    # Hosted image only (NEVER opened locally)
    image_url = "https://raw.githubusercontent.com/Abraham461/EmailBody/main/EmailBody.jpg"

    contact_info = """<div style='text-align:left;'><br>
        Warm Regards,<br>
        Customer Care & Complaints Management<br>
        Operation Department<br><br>
        Phone: +95 9791232222<br>
        Email: customercare@alife.com.mm<br><br>
        A Life Insurance Company Limited<br>
        3rd Floor (A), No. (108), Corner of<br>
        Kabaraye Pagoda Road and Nat Mauk Road,<br>
        Bo Cho (1) Quarter, Bahan Township, Yangon, Myanmar 12201<br>
    </div>"""

    html_body = f"""
    <html>
        <body>
            <img src="{image_url}" style="max-width:100%;" alt="Header"><br>
            <p>{body_text}</p>
            {contact_info}
        </body>
    </html>
    """

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = smtp_user
    msg['To'] = to_email
    msg.set_content(body_text or "Please view this email in HTML format.")
    msg.add_alternative(html_body, subtype='html')

    # Attach Redemption.jpg only (if exists)
    redemption_path = os.path.join('static', 'Redemption.jpg')
    if os.path.exists(redemption_path) and os.path.basename(redemption_path).lower() != "emailbody.jpg":
        with open(redemption_path, 'rb') as f:
            msg.add_attachment(
                f.read(),
                maintype='image',
                subtype='jpeg',
                filename='Redemption.jpg'
            )

    # Optional attachment (block EmailBody.jpg explicitly)
    if attachment_path and os.path.exists(attachment_path):
        if os.path.basename(attachment_path).lower() != "emailbody.jpg":
            with open(attachment_path, 'rb') as f:
                mime_type, _ = mimetypes.guess_type(attachment_path)
                maintype, subtype = mime_type.split('/') if mime_type else ('application', 'octet-stream')
                msg.add_attachment(
                    f.read(),
                    maintype=maintype,
                    subtype=subtype,
                    filename=os.path.basename(attachment_path)
                )
        else:
            logging.warning("Skipped attaching EmailBody.jpg explicitly.")

    # Send email
    try:
        with smtplib.SMTP(smtp_server, smtp_port, timeout=10) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)
        logging.info(f"Email sent to {to_email}")
    except Exception as e:
        logging.error(f"SMTP send failed: {e}")

def generate_cards_from_df(df, output_folder):
    font_label = load_font(FONT_PATH, FONT_SIZE_LABEL)
    font_policy_no = load_font(FONT_PATH, FONT_SIZE_POLICY_NO)
    font_date = load_font(FONT_PATH, FONT_SIZE_DATE)
    font_name = load_font(FONT_PATH, FONT_SIZE_NAME)

    for _, row in df.iterrows():
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
            draw.text((VALID_UNTIL_LABEL_POS[0], VALID_UNTIL_LABEL_POS[1] + bbox[3] - bbox[1] + 5),
                      f"UNTIL - {date}", font=font_date, fill=WHITE)
            draw.text(NAME_POS, name, font=font_name, fill=WHITE)

            filename = os.path.join(output_folder, f"{sanitize_filename(name)}_{sanitize_filename(Card)}.png")
            card.save(filename, format='PNG')

            if email:
                subject = f"Your A-Member Card Awaits You"
                send_email_with_attachment(email, subject, "", filename)

# ---- Other Functions and Flask Routes ----

# Ensure the rest of your Flask routes and other functions are defined here

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    logging.info(f"🚀 Starting Flask app on port {port}")
    app.run(host='0.0.0.0', port=port, threaded=True)

import os
import shutil
import zipfile
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import tempfile
import re

app = Flask(__name__)
app.secret_key = 'supersecretkey'
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'web_cards'
TEMPLATE_IMAGE = 'Card.jpg'
FONT_PATH = 'arial.ttf'
LEFT_MARGIN = 60
POLICY_NO_POS = (LEFT_MARGIN, 250)
VALID_UNTIL_LABEL_POS = (LEFT_MARGIN, 300)
NAME_POS = (LEFT_MARGIN, 420)
FONT_SIZE_POLICY_NO = 36.4
FONT_SIZE_DATE = 18
FONT_SIZE_NAME = 21.9
FONT_SIZE_LABEL = 18
WHITE = (255, 255, 255)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs('static', exist_ok=True)

# Create sample CSV file if it doesn't exist
SAMPLE_CSV_PATH = 'static/sample_cards.csv'
if not os.path.exists(SAMPLE_CSV_PATH):
    sample_data = "Name,Card ID,Date\nJohn Doe,STE 12345 690 7890,2024-12-31\nJane Smith,CII 98765 432 1098,2025-01-15"
    with open(SAMPLE_CSV_PATH, 'w') as f:
        f.write(sample_data)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'csv'}

def format_card_id(card_id):
    """Format card ID while preserving letters and existing spaces"""
    # Remove any non-alphanumeric characters except spaces
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', '', str(card_id))
    # Preserve existing spacing but remove multiple spaces
    cleaned = ' '.join(cleaned.split())
    return cleaned

def load_font(path, size):
    try:
        return ImageFont.truetype(path, int(size))
    except IOError:
        return ImageFont.load_default()

def sanitize_filename(name):
    """Sanitize filenames to remove invalid characters"""
    # Replace spaces with underscores
    name = name.replace(' ', '_')
    # Remove invalid characters
    name = re.sub(r'[\\/*?:"<>|\n]', '', name)
    # Truncate long names
    return name[:100]

def generate_cards_from_df(df, output_folder):
    with Image.open(TEMPLATE_IMAGE) as im:
        TEMPLATE_SIZE = im.size
    font_label = load_font(FONT_PATH, int(FONT_SIZE_LABEL))
    font_policy_no = load_font(FONT_PATH, int(FONT_SIZE_POLICY_NO))
    font_date = load_font(FONT_PATH, int(FONT_SIZE_DATE))
    font_name = load_font(FONT_PATH, int(FONT_SIZE_NAME))
    
    # Use the first three columns regardless of their names
    name_col = df.columns[0]
    card_id_col = df.columns[1]
    date_col = df.columns[2]
    
    for i, row in enumerate(df.iterrows()):
        _, row = row
        # Handle NaN values and ensure strings
        name = str(row[name_col]) if not pd.isna(row[name_col]) else "Unknown"
        card_id = str(row[card_id_col]) if not pd.isna(row[card_id_col]) else "Unknown"
        date = str(row[date_col]) if not pd.isna(row[date_col]) else ""
        
        card = Image.open(TEMPLATE_IMAGE).convert('RGB').resize(TEMPLATE_SIZE)
        draw = ImageDraw.Draw(card)
        
        # Format the card ID while preserving letters and spaces
        formatted_card_id = format_card_id(card_id)
        draw.text(POLICY_NO_POS, formatted_card_id, font=font_policy_no, fill=WHITE)
        
        valid_until_label = "Valid Until"
        draw.text(VALID_UNTIL_LABEL_POS, valid_until_label, font=font_label, fill=WHITE)
        bbox = draw.textbbox(VALID_UNTIL_LABEL_POS, valid_until_label, font=font_label)
        label_width = bbox[2] - bbox[0]
        date_x = VALID_UNTIL_LABEL_POS[0] + label_width + 10
        draw.text((date_x, VALID_UNTIL_LABEL_POS[1]), date, font=font_date, fill=WHITE)
        
        draw.text(NAME_POS, name, font=font_name, fill=WHITE)
        
        # Sanitize filename components
        safe_name = sanitize_filename(name)
        safe_card_id = sanitize_filename(card_id)
        filename = os.path.join(output_folder, f"{safe_name}_{safe_card_id}.png")
        
        try:
            card.save(filename, format='PNG')
        except Exception as e:
            print(f"Error saving {filename}: {str(e)}")
            # Use a fallback filename if there's an issue
            fallback_filename = os.path.join(output_folder, f"card_{i}.png")
            card.save(fallback_filename, format='PNG')

def zip_folder(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.png'):
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, arcname=file)

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
            
            # Check if there are at least 3 columns
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
            flash('Allowed file types: .xlsx, .csv')
            return redirect(request.url)
    
    sample_data = [
        {'Name': 'John Doe', 'Card ID': 'STE 12345 690 7890', 'Date': '2024-12-31'},
        {'Name': 'Jane Smith', 'Card ID': 'CII 98765 432 1098', 'Date': '2025-01-15'}
    ]
    return render_template('index.html', sample_data=sample_data, show_table=True)

@app.route('/download_template')
def download_template():
    return send_file(SAMPLE_CSV_PATH, as_attachment=True, download_name='sample_cards.csv')

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
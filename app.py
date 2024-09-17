import os
import pandas as pd
import requests
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import threading

app = Flask(__name__)

UPLOAD_FOLDER = './uploads'
DOWNLOAD_FOLDER = './downloads'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def find_image_column(df):
    """
    Tries to find the column containing image links.
    If no column with 'image' is found, it also checks for 'background'.
    """
    for col in df.columns:
        if 'image' in str(col).lower() or 'background' in str(col).lower():
            return col
    return None

def download_images(image_links, folder_name, callback):
    folder_path = os.path.join(DOWNLOAD_FOLDER, folder_name)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    total_images = len(image_links)
    for idx, link in enumerate(image_links):
        try:
            response = requests.get(link)
            with open(os.path.join(folder_path, f'image_{idx + 1}.jpg'), 'wb') as file:
                file.write(response.content)
            # Progress callback
            callback(idx + 1, total_images)
        except Exception as e:
            print(f"Error downloading image {link}: {e}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file.save(os.path.join(UPLOAD_FOLDER, filename))
        
        # Read the Excel file and get the sheet names
        excel_path = os.path.join(UPLOAD_FOLDER, filename)
        df = pd.ExcelFile(excel_path)
        sheet_names = df.sheet_names
        
        return jsonify({'sheets': sheet_names, 'file': filename})
    return jsonify({'error': 'File type not allowed'}), 400

@app.route('/download', methods=['POST'])
def download_images_from_tab():
    data = request.json
    file_name = data.get('file_name')
    sheet_name = data.get('sheet_name')
    folder_name = sheet_name  # Set folder name to sheet name
    user_provided_column = data.get('user_column', None)
    
    if not file_name or not sheet_name:
        return jsonify({'error': 'Missing parameters'}), 400
    
    file_path = os.path.join(UPLOAD_FOLDER, file_name)
    
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 400
    
    # Read the selected sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Try to find the column with the image links
    image_column = find_image_column(df)
    if image_column is None:
        if user_provided_column:
            image_column = user_provided_column
        else:
            return jsonify({'error': 'No column containing image links found', 'ask_column': True}), 400
    
    # Extract image links from the found column
    image_links = df[image_column].dropna().tolist()
    
    # Threaded image download to avoid blocking the main process
    def update_progress(current, total):
        # Send progress updates to the client
        print(f'{current}/{total} images downloaded')

    download_thread = threading.Thread(target=download_images, args=(image_links, folder_name, update_progress))
    download_thread.start()

    return jsonify({'message': 'Download started', 'folder_name': folder_name})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))

import os
from pathlib import Path
import time
import shutil
import uuid
from typing import List
from quart import Quart, render_template, request, flash, redirect, url_for, send_file
from werkzeug.utils import secure_filename
from qrcode.constants import ERROR_CORRECT_L
from vcard_qr_generator import main

app = Quart(__name__)


# Check if file extension is allowed
def allowed_file(filename):
    allowed_extensions = {'csv', 'xlsx', 'xls'}
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in allowed_extensions


@app.route('/', methods=['GET'])
async def index():
    app.add_background_task(storage_cleanup)
    return await render_template('index.html')


@app.route('/qr-code', methods=['POST'])
async def upload_file():
    files = await request.files

    # Check if a file was uploaded
    if 'file' not in files:
        await flash('Excel or CSV not uploaded')
        return await redirect(url_for('index'))

    file = files['file']

    # Check if filename is empty
    if file.filename == '':
        await flash('Selected file is invalid (#E0293F)')
        return await redirect(url_for('index'))

    # Get other form data
    form = await request.form
    size = form.get('size') or 4
    border = form.get('border') or 4

    # Validate file and form data
    if not file or not allowed_file(file.filename):
        await flash('Invalid file type. Please upload CSV or Excel file')
        return await redirect(url_for('index'))

    # save the file  locally
    temp_filename = f"{round(time.time() * 1000)}_{secure_filename(file.filename)}"
    sheet_filepath = os.path.join(storage_paths()['uploads'], temp_filename)
    await file.save(sheet_filepath)

    # Generate the vcard QR Codes
    qrcode_options = {
        "version": None,  # Let it be auto-size based on data
        # About 7% or less errors can be corrected. Low error correction means more data capacity
        "error_correction": ERROR_CORRECT_L,

        # how many pixels each “box” of the QR code will have e.g., 10 = ~300x300 pixels (adjust between 1-40)
        "pixel_box_size": size,
        "border": border,  # standard minimum border width is 4
        "color": "black",
        "background_color": "white",

    }

    try:
        qrcode_filename = main(sheet_filepath, storage_paths()[
                               'public'], qrcode_options)
        return redirect(url_for('download_file', filename=qrcode_filename))
    except:
        await flash(f'QR Code generation failed, please try again.')
        return redirect(url_for('index'))


@app.route('/download/<path:filename>', methods=['GET'])
async def download_file(filename: str):

    file_path = os.path.join(
        storage_paths()['public'], secure_filename(filename))

    if not os.path.exists(file_path):
        await flash(f'File {filename} not found')
        return redirect(url_for('index'))

    # Send the file as an attachment
    return await send_file(file_path, as_attachment=True)


def delete_old_files(base_path: str, threshold_time: int, excluded_files: List[str]):
    cut_off_time = time.time() - threshold_time

    for root, dirs, files in os.walk(base_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_mod_time = os.path.getmtime(file_path)
            try:
                # Check if the file is older than set time above
                if file_path not in excluded_files and file_mod_time < cut_off_time:
                    # Delete the file
                    os.remove(file_path)
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")

        # process directories
        for dir_name in dirs:
            dir_path = os.path.join(root, dir_name)
            try:
                # Get directory modification time
                dir_mtime = os.path.getmtime(dir_path)
                if dir_mtime < cut_off_time:
                    shutil.rmtree(dir_path)
            except Exception as e:
                print(f"Error deleting directory {dir_path}: {e}")


# storage_cleanup: cleanup old files in storage drive.
def storage_cleanup(threshold_time=3600):
    # files to exclude
    excluded_files = [
        os.path.join(storage_paths()['public'], 'sample.xlsx')
    ]

    for key, path in storage_paths().items():
        delete_old_files(path, threshold_time, excluded_files)


def storage_paths():
    return {
        "public": "storage/public",
        "uploads": "storage/uploads"
    }


def prepare_drive():
    for _, value in storage_paths().items():
        path = Path(value)
        path.mkdir(exist_ok=True, parents=True)



# ooh yeah!
app.secret_key= os.getenv("SESSION_SECRET", str(uuid.uuid4()))
prepare_drive()

if __name__ == '__main__':
    app.run(debug=os.getenv("DEBUG", False),
    port=os.getenv('PORT', 8000), host=os.getenv("HOST", '0.0.0.0'))

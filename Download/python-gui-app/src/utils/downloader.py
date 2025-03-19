import os
import requests
import zipfile
from tqdm import tqdm

def download_file(url, dest_path, progress_callback=None):
    response = requests.get(url, stream=True)
    total_size = int(response.headers.get('content-length', 0))  # Default to 0 if header is missing
    downloaded_size = 0

    with open(dest_path, 'wb') as file:
        for chunk in response.iter_content(chunk_size=1024):
            if chunk:
                file.write(chunk)
                downloaded_size += len(chunk)
                if progress_callback:
                    progress_callback(downloaded_size, total_size)

def extract_zip(zip_path, extract_to, progress_callback=None):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        total_files = len(zip_ref.namelist())
        for index, file in enumerate(zip_ref.namelist()):
            zip_ref.extract(file, extract_to)
            if progress_callback:
                progress_callback(index + 1, total_files)
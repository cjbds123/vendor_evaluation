import os
import secrets

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or secrets.token_hex(32)
    # Allow override via env var (used in Docker to point at a named volume)
    _db_path = os.environ.get('CPMS_DB_PATH') or os.path.join(BASE_DIR, 'instance', 'cpms.db')
    SQLALCHEMY_DATABASE_URI = 'sqlite:///' + _db_path
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    UPLOAD_FOLDER = os.environ.get('CPMS_UPLOAD_FOLDER') or os.path.join(BASE_DIR, 'uploads')
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max upload
    ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'tiff',
                           'mp4', 'webm', 'mov', 'avi', 'mp3', 'wav', 'ogg',
                           'doc', 'docx', 'xls', 'xlsx', 'pptx', 'zip', 'txt', 'csv'}

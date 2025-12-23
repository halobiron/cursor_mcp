import os
import docker
from minio import Minio
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Docker Client
client = docker.from_env()

# MinIO Config
MINIO_ENDPOINT = os.getenv("MINIO_ENDPOINT", "127.0.0.1:9000")
MINIO_ACCESS_KEY = os.getenv("MINIO_ACCESS_KEY", "minioadmin")
MINIO_SECRET_KEY = os.getenv("MINIO_SECRET_KEY", "minioadmin")
MINIO_SECURE = os.getenv("MINIO_SECURE", "False").lower() == "true"
BUCKET_NAME = os.getenv("MINIO_BUCKET_NAME", "python-outputs")

MINIO_CLIENT = Minio(
    MINIO_ENDPOINT,
    access_key=MINIO_ACCESS_KEY,
    secret_key=MINIO_SECRET_KEY,
    secure=MINIO_SECURE
)

# Workspace Config
HOST_WORKSPACE_DIR = os.path.join(os.getcwd(), "mcp_workspace")
if not os.path.exists(HOST_WORKSPACE_DIR):
    os.makedirs(HOST_WORKSPACE_DIR)

# Initialize Bucket
def ensure_bucket_exists():
    if not MINIO_CLIENT.bucket_exists(BUCKET_NAME):
        MINIO_CLIENT.make_bucket(BUCKET_NAME)

ensure_bucket_exists()

import os
import docker
from minio import Minio
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Docker Client - Tự động tìm Docker socket
def get_docker_client():
    """Tạo Docker client, tự động tìm socket (Desktop hoặc standard)"""
    if not os.getenv("DOCKER_HOST"):
        for sock in ["/var/run/docker.sock", os.path.expanduser("~/.docker/desktop/docker.sock"), "/run/docker.sock"]:
            if os.path.exists(sock):
                os.environ["DOCKER_HOST"] = f"unix://{sock}"
                break
    try:
        return docker.from_env()
    except docker.errors.DockerException as e:
        raise docker.errors.DockerException(f"Không thể kết nối với Docker daemon. Vui lòng đảm bảo Docker đang chạy. Lỗi: {str(e)}")

client = get_docker_client()

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

from mcp.server.fastmcp import FastMCP
import docker
import tarfile
import io
import os
import uuid
from datetime import timedelta
from minio import Minio
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Khởi tạo
mcp = FastMCP("Python Sandbox with Storage")
client = docker.from_env()

# Cấu hình MinIO từ environment variables
MINIO_ENDPOINT = os.getenv("MINIO_ENDPOINT", "localhost:9000")
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

# Tạo bucket nếu chưa có
if not MINIO_CLIENT.bucket_exists(BUCKET_NAME):
    MINIO_CLIENT.make_bucket(BUCKET_NAME)

# Lấy đường dẫn tuyệt đối đến thư mục workspace trên máy thật
# Đây là nơi file sẽ được lưu lại giữa các lần chạy
HOST_WORKSPACE_DIR = os.path.join(os.getcwd(), "mcp_workspace")
if not os.path.exists(HOST_WORKSPACE_DIR):
    os.makedirs(HOST_WORKSPACE_DIR)

@mcp.tool()
def execute_python_code(code: str) -> str:
    """
    Thực thi code Python. 
    Các file cần lưu trữ hoặc đọc lại HÃY để trong thư mục '/app/data'.
    Ví dụ: df.to_csv('/app/data/result.csv')
    """
    container = None
    run_id = str(uuid.uuid4())[:8]
    run_dir = os.path.join(HOST_WORKSPACE_DIR, run_id)
    os.makedirs(run_dir, exist_ok=True)

    try:
        # Cấu hình Volume: Map thư mục riêng của lần chạy này vào container
        volume_config = {
            run_dir: {'bind': '/app/data', 'mode': 'rw'}
        }

        # Chạy container với Volume được mount
        container = client.containers.run(
            "python-sandbox",
            command="python /app/main.py",
            detach=True,
            network_mode="none",
            mem_limit="512m",
            volumes=volume_config, # <--- ĐIỂM QUAN TRỌNG NHẤT
            working_dir="/app/data",
            tty=True
        )

        # (Phần nạp code giống như trước)
        encoded_code = code.encode('utf-8')
        tar_stream = io.BytesIO()
        with tarfile.open(fileobj=tar_stream, mode='w') as tar:
            tar_info = tarfile.TarInfo(name='main.py')
            tar_info.size = len(encoded_code)
            tar.addfile(tar_info, io.BytesIO(encoded_code))
        tar_stream.seek(0)
        container.put_archive("/app", tar_stream)

        # Chờ kết quả
        exit_code = container.wait(timeout=30)
        logs = container.logs().decode("utf-8")
        
        # Upload các file sinh ra lên MinIO
        minio_paths = []
        for filename in os.listdir(run_dir):
            file_path = os.path.join(run_dir, filename)
            if os.path.isfile(file_path):
                object_name = f"{run_id}/{filename}"
                MINIO_CLIENT.fput_object(BUCKET_NAME, object_name, file_path)
                
                # Tạo URL tải về (có hiệu lực trong 7 ngày)
                download_url = MINIO_CLIENT.presigned_get_object(
                    BUCKET_NAME, 
                    object_name
                )
                minio_paths.append(download_url)

        return f"Exit Code: {exit_code['StatusCode']}\nOutput:\n{logs}\n\nDownload Links: {minio_paths}"

    except Exception as e:
        return f"Error: {str(e)}"
    
    finally:
        if container:
            try:
                container.remove(force=True)
            except:
                pass

if __name__ == "__main__":
    mcp.run()

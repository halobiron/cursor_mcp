from mcp.server.fastmcp import FastMCP
import docker
import tarfile
import io
import os
import uuid
from minio import Minio
from dotenv import load_dotenv
import requests

# Load environment variables from .env file
load_dotenv()

# Khởi tạo
mcp = FastMCP("Python Sandbox with Storage")
client = docker.from_env()

# Cấu hình MinIO từ environment variables
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

# Tạo bucket nếu chưa có
if not MINIO_CLIENT.bucket_exists(BUCKET_NAME):
    MINIO_CLIENT.make_bucket(BUCKET_NAME)

# Lấy đường dẫn tuyệt đối đến thư mục workspace trên máy thật
# Đây là nơi file sẽ được lưu lại giữa các lần chạy
HOST_WORKSPACE_DIR = os.path.join(os.getcwd(), "mcp_workspace")
if not os.path.exists(HOST_WORKSPACE_DIR):
    os.makedirs(HOST_WORKSPACE_DIR)

@mcp.tool()
def execute_python_code(code: str, session_id: str = None) -> str:
    """
    Thực thi code Python. 
    
    QUAN TRỌNG: 
    1. Để dùng chung dữ liệu giữa các lần gọi (ví dụ: đọc file đã tải), bạn PHẢI truyền cùng một 'session_id'.
    2. Các file cần lưu trữ hoặc đọc lại HÃY để trong thư mục '/app/data'. Ví dụ: df.to_csv('/app/data/result.csv')
    3. Nếu bạn chưa có 'session_id', hãy để trống ở lần gọi đầu tiên, hệ thống sẽ cấp cho bạn một ID.
    
    Args:
        code: Code Python cần chạy.
        session_id: ID của phiên làm việc (Session ID). Hãy tái sử dụng ID từ các lần gọi trước.
    """
    container = None
    if not session_id:
        session_id = str(uuid.uuid4())[:8]
    session_dir = os.path.join(HOST_WORKSPACE_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    try:
        # Cấu hình Volume: Map thư mục riêng của lần chạy này vào container
        volume_config = {
            session_dir: {'bind': '/app/data', 'mode': 'rw'}
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
        for filename in os.listdir(session_dir):
            file_path = os.path.join(session_dir, filename)
            if os.path.isfile(file_path) and filename != "document":
                object_name = f"{session_id}/{filename}"
                MINIO_CLIENT.fput_object(BUCKET_NAME, object_name, file_path)
                
                # Tạo URL tải về (có hiệu lực trong 7 ngày)
                download_url = MINIO_CLIENT.presigned_get_object(
                    BUCKET_NAME, 
                    object_name
                )
                minio_paths.append({"filename": filename, "url": download_url})

        return f"[Session ID: {session_id}]\nExit Code: {exit_code['StatusCode']}\nOutput:\n{logs}\n\nDownload Links: {minio_paths}"

    except Exception as e:
        return f"Error (Session ID: {session_id}): {str(e)}"
    
    finally:
        if container:
            try:
                container.remove(force=True)
            except:
                pass

@mcp.tool()
def download_document(document_url: str, filename: str = "document", session_id: str = None) -> str:
    """
    Tải tài liệu từ link và lưu vào thư mục dữ liệu để xử lý sau.
    
    QUAN TRỌNG:
    1. Tài liệu sẽ được lưu tại '/app/data/{filename}'.
    2. Bạn PHẢI sử dụng 'session_id' được trả về trong các lần gọi 'execute_python_code' tiếp theo để truy cập file này.
    
    Args:
        document_url: URL của tài liệu (CSV, TXT, JSON, Excel, v.v.)
        filename: Tên file muốn lưu (mặc định là 'document').
        session_id: (Tùy chọn) ID của phiên làm việc. Nếu có sẵn, hãy truyền vào để lưu chung chỗ.
    """
    if not session_id:
        session_id = str(uuid.uuid4())[:8]
    session_dir = os.path.join(HOST_WORKSPACE_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    try:
        response = requests.get(document_url, timeout=30)
        response.raise_for_status()
        
        # Lưu file
        document_path = os.path.join(session_dir, filename)
        with open(document_path, "wb") as f:
            f.write(response.content)
        
        return f"Tải thành công! File đã lưu tại '/app/data/{filename}'.\n\nQUAN TRỌNG: Hãy sử dụng session_id='{session_id}' cho các lần gọi tiếp theo."

    except Exception as e:
        return f"Lỗi tải file (Session ID: {session_id}): {str(e)}"

if __name__ == "__main__":
    mcp.run()

import os
import io
import uuid
import tarfile
from .config import client, MINIO_CLIENT, BUCKET_NAME, HOST_WORKSPACE_DIR

def execute_python_code(code: str, session_id: str = None) -> str:
    """Thực thi code Python."""
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

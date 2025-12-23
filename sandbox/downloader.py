import os
import uuid
import requests
from .config import HOST_WORKSPACE_DIR

def download_document(document_url: str, filename: str = "document", session_id: str = None) -> str:
    """Tải tài liệu từ link."""
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

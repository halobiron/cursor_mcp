# Cập nhật
## 19.12.2
- **Quản lý phiên:** Hỗ trợ `session_id` để duy trì dữ liệu giữa các lần gọi tool.
- **Công cụ tải tài liệu:** Thêm tool `download_document` hỗ trợ tải file từ URL với tùy chọn đặt tên file (`filename`).
- **Cải tiến phản hồi:** Link tải từ MinIO giờ bao gồm cả tên file để dễ nhận diện.

## 19.12.1
- **Cô lập dữ liệu:** Mỗi phiên chạy dùng thư mục UUID riêng, tránh xung đột.
- **MinIO:** Tự động upload & tạo link tải (Presigned URL 7 ngày).
- **Tối ưu:** `working_dir` tại `/app/data` giúp thao tác file nhanh gọn và ít lỗi.

## Cài đặt MinIO

```bash
docker run -d \
  -p 9000:9000 \
  -p 9001:9001 \
  --name minio \
  -e "MINIO_ROOT_USER=minioadmin" \
  -e "MINIO_ROOT_PASSWORD=minioadmin" \
  minio/minio server /data --console-address ":9001"
```
FROM python:3.13-slim

WORKDIR /app

# Cài đặt các package cần thiết
RUN pip install --no-cache-dir pandas numpy requests minio openpyxl

# Tạo thư mục data để mount volume
RUN mkdir -p /app/data

WORKDIR /app/data

# Đợi cho đến khi file main.py xuất hiện rồi mới chạy lệnh
ENTRYPOINT ["sh", "-c", "until [ -f /app/main.py ]; do sleep 0.1; done; exec \"$@\"", "--"]

CMD ["python", "/app/main.py"]
FROM python:3.12-slim
WORKDIR /app
RUN pip install --no-cache-dir openpyxl
COPY index.html upload.html server.py update_data.py graph_sync.py /app/
ENV HPS_DATA_DIR=/data \
    PYTHONUNBUFFERED=1 \
    TZ=Asia/Jakarta
EXPOSE 8080
VOLUME ["/data"]
HEALTHCHECK --interval=60s --timeout=10s --start-period=20s \
  CMD python -c "import urllib.request;urllib.request.urlopen('http://127.0.0.1:8080/api/health',timeout=5)"
CMD ["python", "server.py", "8080"]

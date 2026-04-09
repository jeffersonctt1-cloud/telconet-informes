FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip3 install --no-cache-dir flask python-docx Pillow gunicorn
COPY . .
CMD gunicorn server:app --bind 0.0.0.0:$PORT --timeout 120

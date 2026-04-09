FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN python -m pip install --upgrade pip && \
    python -m pip install flask python-docx Pillow gunicorn
CMD gunicorn server:app --bind 0.0.0.0:$PORT --timeout 120

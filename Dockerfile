FROM python:3.11-slim

WORKDIR /app

# Install system dependencies for reportlab/matplotlib
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    libfreetype6-dev \
    libpng-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .
COPY templates/index.html ./templates/

EXPOSE 5000

CMD ["python", "app.py"]

# ── Build stage ───────────────────────────────────────────────
FROM python:3.11-slim AS base

# Keeps Python from generating .pyc files and enables real-time logs
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Install system deps
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Install Python deps first (layer cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY app.py .
COPY index.html .

# Create uploads directory
RUN mkdir -p uploads

# Expose port (Render injects $PORT at runtime; default 5000 for local)
EXPOSE 5000

# Gunicorn: 2 workers, 120s timeout for large uploads
CMD gunicorn app:app \
    --bind 0.0.0.0:${PORT:-5000} \
    --workers 2 \
    --timeout 120 \
    --access-logfile - \
    --error-logfile -

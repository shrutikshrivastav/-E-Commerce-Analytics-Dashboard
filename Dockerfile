# Use a lightweight Python image
FROM python:3.9-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y gcc

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY app.py .
COPY static /app/static  # Assuming you move the HTML/JS to a static folder in the real setup
# For the single file demo, we might just serve the template

# Expose port
EXPOSE 5000

# Run Gunicorn
CMD ["gunicorn", "-b", "0.0.0.0:5000", "app:app"]

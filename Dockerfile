# Sử dụng Python 3.11 slim image
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Copy requirements và install dependencies
COPY requirements_firebase.txt .
RUN pip install --no-cache-dir -r requirements_firebase.txt

# Copy application code
COPY app_firebase.py .
COPY filter.py .
COPY cas_database.json .
COPY index.html .
COPY sign_in.html .

# Tạo thư mục uploads và results
RUN mkdir -p uploads results

# Expose port
EXPOSE 8080

# Set environment variables
ENV PORT=8080
ENV PYTHONUNBUFFERED=1

# Run with gunicorn
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 app_firebase:app

# Use official Python slim image with fixed version
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Copy requirements and install pinned versions only
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy rest of your app code
COPY . .

# Expose port (default Flask port)
EXPOSE 5000

# Run your flask app
CMD ["python", "app.py"]

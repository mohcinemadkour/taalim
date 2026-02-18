FROM python:3.11-slim

# Install Chromium and dependencies for Kaleido
RUN apt-get update && apt-get install -y \
    chromium \
    chromium-driver \
    fonts-liberation \
    fonts-noto-cjk \
    fonts-noto-color-emoji \
    fonts-noto-ui-arabic \
    fonts-arabeyes \
    && rm -rf /var/lib/apt/lists/*

# Set environment variable for Kaleido to find Chromium
ENV CHROMIUM_PATH=/usr/bin/chromium
ENV KALEIDO_CHROMIUM_EXECUTABLE=/usr/bin/chromium

WORKDIR /app

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app
COPY . .

# Expose port
EXPOSE 10000

# Run Streamlit
CMD ["streamlit", "run", "app.py", "--server.port=10000", "--server.address=0.0.0.0"]

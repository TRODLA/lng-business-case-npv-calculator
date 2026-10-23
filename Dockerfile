FROM python:3.12-slim

# Create non-root user with fixed UID
RUN useradd -u 1001 -m appuser

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Set correct ownership
RUN chown -R 1001:1001 /app

# ✅ Use numeric UID (THIS IS THE FIX)
USER 1001

EXPOSE 8501

CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0", "--server.port=8501"]

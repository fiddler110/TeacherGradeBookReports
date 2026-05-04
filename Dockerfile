# ── Build stage ───────────────────────────────────────────────────────────────
FROM python:3.12-slim AS base

# Install system dependencies needed by reportlab (fonts, etc.)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libfreetype6 \
    libfontconfig1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies first (layer-cached unless requirements change)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY generate_reports.py .
COPY app.py .
COPY templates/ templates/
COPY static/ static/

# Temp directory for uploaded files and generated reports
ENV UPLOAD_ROOT=/tmp/reportgen

# SQLite database for user accounts (mount a volume here for persistence)
ENV DB_PATH=/data/users.db
VOLUME ["/data"]

EXPOSE 8000

# 2 workers is appropriate for a small single-tenant tool;
# increase via the GUNICORN_WORKERS env var if needed
ENV GUNICORN_WORKERS=2

CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:8000 --workers ${GUNICORN_WORKERS:-2} --timeout 120 --access-logfile - app:app"]

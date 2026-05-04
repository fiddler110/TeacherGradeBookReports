# ── Builder stage ────────────────────────────────────────────────────────────
# Installs all Python dependencies into an isolated venv.
# pip, setuptools, wheel and any build-time headers stay here and never
# make it into the final image.
FROM python:3.12-slim AS builder

WORKDIR /build

# Create the venv that will be copied to the runtime stage
RUN python -m venv /venv
ENV PATH="/venv/bin:$PATH"

COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# ── Runtime stage ─────────────────────────────────────────────────────────────
# Starts from the same slim base but contains only:
#   • runtime system libs (no dev headers, no build tools)
#   • the pre-built venv copied from builder
#   • application source
#   • a non-root service account
FROM python:3.12-slim AS runtime

# Runtime-only system libraries required by ReportLab / Pillow
RUN apt-get update && apt-get install -y --no-install-recommends \
    libfreetype6 \
    libfontconfig1 \
    && rm -rf /var/lib/apt/lists/*

# Non-root service account – never run as root inside the container
RUN groupadd --gid 1001 appgroup \
    && useradd --uid 1001 --gid appgroup --no-create-home --shell /bin/false appuser

# Copy the pre-built virtual environment from the builder stage
COPY --from=builder /venv /venv

WORKDIR /app

# Copy only the files the application needs at runtime
COPY generate_reports.py app.py ./
COPY templates/ templates/
COPY static/ static/

# Pre-compile source to .pyc – reduces first-request latency and avoids
# needing a writable source directory at runtime.
# Also create /data with correct ownership before switching user.
RUN python -m compileall -q . \
    && chown -R appuser:appgroup /app \
    && mkdir -p /data && chown appuser:appgroup /data

ENV PATH="/venv/bin:$PATH" \
    # Prevent Python from writing .pyc files at runtime (already compiled above)
    PYTHONDONTWRITEBYTECODE=1 \
    # Ensure stdout/stderr are flushed immediately (important for log visibility)
    PYTHONUNBUFFERED=1 \
    UPLOAD_ROOT=/tmp/reportgen \
    DB_PATH=/data/users.db \
    GUNICORN_WORKERS=2

# SQLite database for user accounts (mount a volume here for persistence)
VOLUME ["/data"]

EXPOSE 8000

# All subsequent processes run as the non-root service account
USER appuser

CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:8000 --workers ${GUNICORN_WORKERS:-2} --timeout 120 --access-logfile - app:app"]

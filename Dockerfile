FROM python:3.11-slim

WORKDIR /app

# Install curl for healthcheck
RUN apt-get update && apt-get install -y --no-install-recommends curl \
    && rm -rf /var/lib/apt/lists/*

# Install dependencies and the bomgen package
COPY pyproject.toml .
COPY README.md .
COPY src/ src/
COPY templates/ templates/

RUN pip install --no-cache-dir -e .

# Ensure bomgen package is importable when running ui.py
ENV PYTHONPATH=/app/src

EXPOSE 8080

# Use file path (not module) - Streamlit recommends this for Docker
# Bind to 0.0.0.0 so the app accepts external connections
HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
    CMD curl -f http://localhost:8080/_stcore/health || exit 1

CMD ["streamlit", "run", "src/bomgen/ui.py", \
    "--server.port=8080", \
    "--server.address=0.0.0.0", \
    "--server.headless=true", \
    "--server.fileWatcherType=none"]

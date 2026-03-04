FROM python:3.11-slim

WORKDIR /app

# Install dependencies and the bomgen package
COPY pyproject.toml .
COPY README.md .
COPY src/ src/

RUN pip install --no-cache-dir -e .

# Run Streamlit as a module (preserves package context for relative imports)
EXPOSE 8080
CMD ["streamlit", "run", "bomgen.ui", "--server.port=8080", "--server.address=0.0.0.0"]

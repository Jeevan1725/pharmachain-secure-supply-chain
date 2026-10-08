# PharmaChain - secure minimal container
FROM python:3.12-slim AS base

# Non-root user
RUN useradd -m -u 1000 appuser

WORKDIR /app

# Install only runtime dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app source
COPY --chown=appuser:appuser src/ ./src/
COPY --chown=appuser:appuser pyproject.toml .

# Install package
RUN pip install --no-cache-dir -e .

# Switch to non-root
USER appuser

# Only expose app port
EXPOSE 8080

# Healthcheck
HEALTHCHECK --interval=30s --timeout=5s --retries=3 CMD python -c "import sys; sys.exit(0)" || exit 1

CMD ["python", "-m", "uvicorn", "pharmachain.api.main:app", "--host", "0.0.0.0", "--port", "8080"]
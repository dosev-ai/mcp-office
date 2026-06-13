FROM python:3.11-slim
WORKDIR /app
COPY wordmcp/ ./wordmcp/
RUN pip install --no-cache-dir -e ./wordmcp
ENV WORD_ENABLE_WRITE=false
ENV WORD_ALLOWLIST_ROOTS=/tmp
CMD ["python", "-m", "wordmcp.server"]

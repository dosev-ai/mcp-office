FROM python:3.11-slim
WORKDIR /app
COPY shared/ ./shared/
COPY pptmcp/ ./pptmcp/
RUN pip install --no-cache-dir -e ./shared && \
    pip install --no-cache-dir -e ./pptmcp
ENV PPT_ENABLE_WRITE=false
ENV PPT_ALLOWLIST_ROOTS=/tmp
CMD ["python", "-m", "pptmcp.server"]

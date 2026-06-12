# Dockerfile for glama.ai MCP server verification
# Runs excelmcp (Excel MCP server) — no Windows/COM required, uses openpyxl

FROM python:3.11-slim

WORKDIR /app

COPY shared/ ./shared/
COPY excelmcp/ ./excelmcp/

RUN pip install --no-cache-dir -e ./shared && \
    pip install --no-cache-dir -e ./excelmcp

ENV EXCEL_ENABLE_WRITE=false
ENV EXCEL_ALLOWLIST_ROOTS=/tmp

CMD ["python", "-m", "excelmcp.server"]

# ppt-academizer API (for Render/Fly/Railway — not Netlify)
FROM python:3.12-slim-bookworm

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libxml2 libxslt1.1 \
    git \
    openssh-client \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY api api
COPY core core
COPY engine engine
COPY scripts scripts
COPY docker/entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

ENV PYTHONPATH=/app/engine:/app
ENV PPT_ACADEMIZER_SKIP_PP_REPAIR=1
ENV PORT=8765
ENV TEMPLATE_PPTX=/tmp/academy-template.pptx

EXPOSE 8765

ENTRYPOINT ["/entrypoint.sh"]
CMD ["python", "-m", "uvicorn", "api.main:app", "--host", "0.0.0.0", "--port", "8765", "--app-dir", "/app"]

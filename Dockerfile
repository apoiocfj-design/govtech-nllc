# Sistema GovTech NLLC - imagem para nuvem (Render / Cloud Run)
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=utf-8 \
    PORT=8080

WORKDIR /app

# Chromium headless para capturar a pagina oficial do PNCP em PDF (opcional; ha fallback em ReportLab)
RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium fonts-dejavu-core ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY backend/ /app/backend/
COPY database/ /app/database/
COPY frontend/ /app/frontend/

# Descompacta o catalogo CNBS durante o build (evita atraso a cada inicializacao)
RUN python -c "import gzip,shutil; shutil.copyfileobj(gzip.open('database/cnbs_catalog.db.gz','rb'), open('database/cnbs_catalog.db','wb'))" \
    && rm -f database/cnbs_catalog.db.gz

EXPOSE 8080
CMD ["python", "backend/main.py"]

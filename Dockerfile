# ==============================================================================
# DOCKERFILE - SISTEMA GOVTECH NLLC (CLOUD RUN / FIREBASE HOSTING)
# Imagem leve baseada em Python 3.12 Slim (Linux Alpine/Debian)
# ==============================================================================
FROM python:3.12-slim

# Evita criação de arquivos .pyc e força saída direta de logs no console
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=8080

WORKDIR /app

# Instala dependências do sistema necessárias para compilação mínima
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Instala dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia código-fonte e banco de dados SQLite indexado (CNBS)
COPY backend/ /app/backend/
COPY database/ /app/database/
COPY frontend/ /app/frontend/

EXPOSE 8080

CMD ["python", "backend/main.py"]

# Imagen base con Chrome y noVNC
FROM chromedriver:stable

# Crear carpetas necesarias e instalar dependencias del sistema (LibreOffice para conversión docx a pdf)
# Debian 11 (Bullseye) pasó a archivo histórico, por lo que usamos archive.debian.org
RUN echo "deb http://archive.debian.org/debian bullseye main" > /etc/apt/sources.list && \
    echo "deb http://archive.debian.org/debian-security bullseye-security main" >> /etc/apt/sources.list && \
    apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

RUN mkdir -p /app/Downloads /app/logs

# Copiar requirements.txt
COPY requirements.txt /app/

# Instalar dependencias Python
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copiar código del proyecto
COPY Codigo/ /app/Codigo

# Copiar supervisord.conf
COPY supervisord.conf /app/

# Workdir
WORKDIR /app

# CMD para levantar supervisord
CMD ["supervisord", "-c", "/app/supervisord.conf"]

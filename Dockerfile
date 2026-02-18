# --------------------------------------------------
# Imagen base con Chrome y noVNC
# --------------------------------------------------
FROM chromedriver:stable

# Crear carpetas necesarias
RUN mkdir -p /app/Downloads

# Copiar requirements.txt
COPY requirements.txt /app/

# Instalar dependencias Python
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copiar supervisord.conf
COPY supervisord.conf /app/

# Asegurar permisos de lectura/ejecución (volúmenes externos no permiten chown)
#RUN chmod -R 755 /app

# Workdir y usuario
WORKDIR /app
USER user1

# CMD para levantar supervisord
CMD ["supervisord", "-c", "/app/supervisord.conf"]

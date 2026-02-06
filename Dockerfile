#Dockerfile con entorno gráfico + VNC + Chrome
FROM python:3.11-bullseye

# Instala dependencias del sistema necesarias
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        wget curl unzip gnupg ca-certificates fonts-liberation \
        libnss3 libx11-xcb1 libxcomposite1 libxdamage1 libxrandr2 \
        libasound2 libatk-bridge2.0-0 libatk1.0-0 libcups2 libdbus-1-3 \
        xdg-utils libgbm1 libgtk-3-0 libu2f-udev libvulkan1 \
        git xvfb x11vnc supervisor net-tools python3-tk \
        scrot libxss1 libxtst6 xdotool wmctrl dbus-x11 fonts-dejavu \
        gnome-screenshot openbox x11-utils libx11-6 libxext6 libxrender1 \
        xauth \
        pandoc \
        wkhtmltopdf \
    && rm -rf /var/lib/apt/lists/*

# Instalar Google Chrome estable
RUN curl -fsSL https://dl.google.com/linux/linux_signing_key.pub | gpg --dearmor -o /usr/share/keyrings/google.gpg \
 && echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google.gpg] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google.list \
 && apt-get update && apt-get install -y google-chrome-stable

RUN apt-get update && apt-get install -y wget \
  && wget https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6.1-2/wkhtmltox_0.12.6.1-2.bullseye_amd64.deb \
  && apt install -y ./wkhtmltox_0.12.6.1-2.bullseye_amd64.deb \
  && rm ./wkhtmltox_0.12.6.1-2.bullseye_amd64.deb

# Instalar noVNC
RUN git clone https://github.com/novnc/noVNC.git /opt/novnc && \
    git clone https://github.com/novnc/websockify /opt/websockify && \
    ln -s /opt/novnc/vnc.html /opt/novnc/index.html

WORKDIR /app

## Copiar proyecto
#COPY . .

# 👉 SOLO dependencias (no el código)
COPY requirements.txt .

# Instalar dependencias python
RUN pip install --no-cache-dir -r requirements.txt

# Crear carpetas necesarias
RUN mkdir -p /app/Downloads

# Copiar supervisord.conf
COPY supervisord.conf /app/supervisord.conf

ENV PYTHONUNBUFFERED=1

CMD ["supervisord", "-c", "/app/supervisord.conf"]

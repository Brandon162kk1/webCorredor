# -- Imports --
import time
import subprocess

def esperar_ventana(nombre, timeout=10):
    inicio = time.time()
    while time.time() - inicio < timeout:
        result = subprocess.run(
            ["xdotool", "search", "--name", nombre],
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL
        )
        if result.stdout.strip():
            return True
        time.sleep(0.5)
    return False

def bloquear_interaccion():
    subprocess.run(["x11vnc", "-remote", "viewonly"], check=False)
    print("🔒 Interacción humana BLOQUEADA (VNC view-only)")

def desbloquear_interaccion():
    subprocess.run(["x11vnc", "-remote", "noviewonly"], check=False)
    print("✋ Interacción humana HABILITADA")
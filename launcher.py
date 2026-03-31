"""
OnepageReport Automação - Launcher
Inicia o servidor e abre o navegador automaticamente.
Duplo clique para executar (ou rode: python launcher.py)
"""

import sys
import os
import socket
import threading
import webbrowser
import time

# ── Garante que os módulos bundled sejam encontrados (PyInstaller) ───────────
if hasattr(sys, "_MEIPASS"):
    sys.path.insert(0, sys._MEIPASS)

# ── Importa a aplicação Flask ────────────────────────────────────────────────
from app import app  # noqa: E402

PORT_PADRAO = 5000


def _porta_livre(porta):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", porta)) != 0


def _encontrar_porta(inicio=PORT_PADRAO):
    for p in range(inicio, inicio + 20):
        if _porta_livre(p):
            return p
    return inicio


def _abrir_navegador(url, delay=1.8):
    time.sleep(delay)
    webbrowser.open(url)


if __name__ == "__main__":
    porta = _encontrar_porta(PORT_PADRAO)
    url   = f"http://localhost:{porta}"

    banner = f"""
╔══════════════════════════════════════════════╗
║        OnepageReport Automação              ║
╠══════════════════════════════════════════════╣
║  Servidor rodando em: {url:<23}║
║                                              ║
║  O navegador abrirá automaticamente.         ║
║  Feche esta janela para encerrar.            ║
╚══════════════════════════════════════════════╝
"""
    print(banner)

    # Abre o navegador em background
    threading.Thread(
        target=_abrir_navegador, args=(url,), daemon=True
    ).start()

    # Inicia o servidor Flask (bloqueia até a janela ser fechada)
    app.run(
        host="127.0.0.1",
        port=porta,
        debug=False,
        use_reloader=False,
    )

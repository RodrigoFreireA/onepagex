"""
OnePageReport - Minimal local HTTP server.

Usage:
    python server.py

Requires Python 3.8+ and project dependencies installed.
"""

import base64
import http.server
import json
import os
import shutil
import socketserver
import sys
import tempfile
import threading
import urllib.parse
import unicodedata
import webbrowser
from datetime import datetime
from email.parser import BytesParser


PORT = 5000
HERE = os.path.dirname(os.path.abspath(__file__))

# Ensure local imports resolve.
sys.path.insert(0, HERE)


def _safe_filename(name):
    """
    Build a safe Content-Disposition filename for latin-1 headers.
    """
    try:
        name.encode("latin-1")
        return f'attachment; filename="{name}"'
    except UnicodeEncodeError:
        ascii_name = (
            unicodedata.normalize("NFKD", name)
            .encode("ascii", "ignore")
            .decode("ascii")
        )
        if not ascii_name:
            ascii_name = "Dashboard.xlsx"
        encoded = urllib.parse.quote(name, safe="")
        return f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{encoded}"


def _encode_warnings_header(warnings):
    raw = json.dumps(warnings, ensure_ascii=False)
    return base64.b64encode(raw.encode("utf-8")).decode("ascii")


def _setup_shutdown_handler():
    """
    Register console shutdown handlers for Windows.
    """
    if sys.platform != "win32":
        return

    import ctypes
    import ctypes.wintypes
    import signal

    def _sig_handler(_sig, _frame):
        print("\n[OK] Servidor encerrado.")
        os._exit(0)

    try:
        signal.signal(signal.SIGINT, _sig_handler)
        signal.signal(signal.SIGTERM, _sig_handler)
        if hasattr(signal, "SIGBREAK"):
            signal.signal(signal.SIGBREAK, _sig_handler)
    except Exception:
        pass

    try:
        HandlerRoutine = ctypes.WINFUNCTYPE(
            ctypes.wintypes.BOOL, ctypes.wintypes.DWORD
        )

        def _console_handler(ctrl_type):
            if ctrl_type in (0, 1, 2, 5, 6):
                print("\n[OK] Janela fechada - servidor encerrado.")
                os._exit(0)
            return False

        handler_ref = HandlerRoutine(_console_handler)
        ctypes.windll.kernel32.SetConsoleCtrlHandler(handler_ref, True)
        global _WIN_CTRL_HANDLER  # noqa: PLW0603
        _WIN_CTRL_HANDLER = handler_ref
    except Exception:
        pass


def _find_template():
    for name in ("template.next.clean.xlsx", "template.next.xlsx", "template.xlsx"):
        path = os.path.join(HERE, name)
        if os.path.exists(path):
            return path
    return None


def _parse_multipart(handler):
    """
    Manual multipart/form-data parser.
    Returns dict where file fields are:
      {"filename": "...", "data": b"..."}
    """
    content_type = handler.headers.get("Content-Type", "")
    content_length = int(handler.headers.get("Content-Length", 0))
    body = handler.rfile.read(content_length)

    boundary = None
    for part in str(content_type).split(";"):
        part = part.strip()
        if part.startswith("boundary="):
            boundary = part[len("boundary="):].strip().strip('"')
            break
    if not boundary:
        return {}

    sep = ("--" + boundary).encode()
    fields = {}

    for chunk in body.split(sep):
        chunk = chunk.strip(b"\r\n")
        if not chunk or chunk == b"--":
            continue
        if chunk.startswith(b"--"):
            break

        if b"\r\n\r\n" in chunk:
            raw_headers, part_body = chunk.split(b"\r\n\r\n", 1)
        elif b"\n\n" in chunk:
            raw_headers, part_body = chunk.split(b"\n\n", 1)
        else:
            continue

        part_body = part_body.rstrip(b"\r\n")
        headers = BytesParser().parsebytes(raw_headers + b"\r\n\r\n")
        disp = str(headers.get("Content-Disposition", ""))

        name = None
        filename = None
        for item in disp.split(";"):
            item = item.strip()
            if item.startswith("name="):
                name = item[5:].strip().strip('"')
            elif item.startswith("filename="):
                filename = item[9:].strip().strip('"')

        if name is None:
            continue

        if filename:
            fields[name] = {"filename": filename, "data": part_body}
        else:
            fields[name] = part_body.decode("utf-8", errors="replace")

    return fields


class DashboardHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, _fmt, *_args):
        # Keep terminal output clean.
        return

    def log_error(self, fmt, *args):
        sys.stderr.write(f"[ERRO servidor] {fmt % args}\n")

    def do_GET(self):
        path = self.path.split("?")[0]

        if path in ("/", "/index.html"):
            self._serve_file(os.path.join(HERE, "index.html"), "text/html; charset=utf-8")
            return

        if path == "/status":
            self._json(200, {"status": "ok", "version": "server.py"})
            return

        if path == "/baixar-starter":
            tpl = _find_template()
            if not tpl:
                self._json(404, {"erro": "Template nao encontrado."})
                return
            with open(tpl, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Disposition", 'attachment; filename="Modelo_Manual.xlsx"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        if path == "/shutdown":
            self._json(200, {"ok": True})
            threading.Thread(target=lambda: os._exit(0), daemon=True).start()
            return

        self._json(404, {"erro": "Rota nao encontrada."})

    def do_POST(self):
        path = self.path.split("?")[0]
        if path == "/gerar":
            self._handle_gerar()
            return
        if path == "/validar-base":
            self._handle_validar_base()
            return
        if path == "/shutdown":
            self._json(200, {"ok": True})
            threading.Thread(target=lambda: os._exit(0), daemon=True).start()
            return
        self._json(404, {"erro": "Rota nao encontrada."})

    def _handle_gerar(self):
        content_type = str(self.headers.get("Content-Type", ""))

        # Mode A: JSON payload already precomputed.
        if "application/json" in content_type:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)
            try:
                payload = json.loads(body)
            except Exception as e:
                self._json(400, {"erro": f"JSON invalido: {e}"})
                return
            self._gerar_from_json(payload)
            return

        # Mode B: multipart upload with base Planner export.
        if "multipart/form-data" not in content_type:
            self._json(400, {"erro": "Esperado multipart/form-data ou application/json."})
            return

        fields = _parse_multipart(self)

        arquivo_field = fields.get("arquivo")
        if not arquivo_field or not isinstance(arquivo_field, dict):
            self._json(400, {"erro": "Nenhum arquivo enviado."})
            return

        filename = arquivo_field.get("filename", "")
        file_bytes = arquivo_field.get("data", b"")
        if not filename.lower().endswith(".xlsx"):
            self._json(400, {"erro": "Arquivo invalido. Envie um .xlsx exportado do Planner."})
            return

        author = fields.get("autor", "Gerado automaticamente")
        if isinstance(author, bytes):
            author = author.decode("utf-8", errors="replace")

        ignore_labels = fields.get("rotulos_ignorar", "")
        if isinstance(ignore_labels, bytes):
            ignore_labels = ignore_labels.decode("utf-8", errors="replace")

        modo = fields.get("modo", "xlsx")
        if isinstance(modo, bytes):
            modo = modo.decode("utf-8", errors="replace")

        dm_field = fields.get("dados_manuais")
        dm_path = None

        form_data = None
        df_field = fields.get("dados_formulario", "")
        if isinstance(df_field, bytes):
            df_field = df_field.decode("utf-8", errors="replace")
        if isinstance(df_field, str) and df_field.strip():
            try:
                form_data = json.loads(df_field)
            except Exception as e:
                self._json(400, {"erro": f"dados_formulario JSON invalido: {e}"})
                return

        work_dir = tempfile.mkdtemp()
        input_path = os.path.join(work_dir, "input.xlsx")
        output_path = os.path.join(work_dir, "dashboard.xlsx")

        try:
            with open(input_path, "wb") as f:
                f.write(file_bytes)

            template_path = _find_template()
            if not template_path:
                self._json(
                    500,
                    {
                        "erro": (
                            "Template nao encontrado. Coloque template.next.clean.xlsx "
                            "(ou template.next.xlsx) na pasta do server.py."
                        )
                    },
                )
                return

            if dm_field and isinstance(dm_field, dict):
                dm_filename = dm_field.get("filename", "")
                dm_bytes = dm_field.get("data", b"")
                if dm_filename.lower().endswith(".xlsx"):
                    dm_path = os.path.join(work_dir, "dados_manuais.xlsx")
                    with open(dm_path, "wb") as f:
                        f.write(dm_bytes)

            from generate_dashboard import compute_json, fill_template

            if modo == "json":
                self._json(200, compute_json(input_path, ignore_labels=ignore_labels))
                return

            fill_template(
                template_path,
                input_path,
                output_path,
                author=author,
                dados_manuais_path=dm_path,
                ignore_labels=ignore_labels,
                form_data=form_data,
            )

            warnings = []
            try:
                warnings = compute_json(
                    input_path, ignore_labels=ignore_labels
                ).get("warnings", [])
            except Exception as warn_e:
                print(f"[aviso] Falha ao coletar warnings: {warn_e}")

            if not os.path.exists(output_path):
                self._json(500, {"erro": "Arquivo de saida nao foi gerado."})
                return

            with open(output_path, "rb") as f:
                xlsx_bytes = f.read()

            ts = datetime.now().strftime("%Y%m%d_%H%M")
            base_name = os.path.splitext(os.path.basename(filename))[0]
            out_name = f"Dashboard_{base_name}_{ts}.xlsx"

            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Disposition", _safe_filename(out_name))
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self.send_header("X-Dashboard-Warnings", _encode_warnings_header(warnings))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header(
                "Access-Control-Expose-Headers",
                "Content-Disposition, X-Dashboard-Warnings",
            )
            self.end_headers()
            self.wfile.write(xlsx_bytes)

            if warnings:
                print(f"[WARN] {len(warnings)} aviso(s) ao gerar {out_name}:")
                for w in warnings:
                    nivel = w.get("nivel", "info").upper()
                    print(f"       [{nivel}] {w.get('msg', '')}")
            else:
                print(f"[OK] Dashboard gerado: {out_name}")

        except Exception as e:
            import traceback

            tb = traceback.format_exc()
            print(f"[ERRO] {e}\n{tb}")

            msg = str(e)
            if "No sheet named" in msg or "Worksheet" in msg:
                msg = (
                    "Aba nao encontrada no template ou arquivo. "
                    f"Verifique se o template e o correto - {e}"
                )
            elif "openpyxl" in msg.lower() or "zipfile" in msg.lower():
                msg = f"Arquivo corrompido ou formato invalido - {e}"
            elif "KeyError" in type(e).__name__:
                msg = (
                    "Coluna esperada nao encontrada no export do Planner. "
                    f"Certifique-se de exportar do Microsoft Planner - {e}"
                )
            elif "date" in msg.lower() or "datetime" in msg.lower():
                msg = (
                    "Erro ao processar datas. Verifique se as colunas de data "
                    f"estao corretas no arquivo - {e}"
                )
            self._json(500, {"erro": msg})
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    def _handle_validar_base(self):
        content_type = str(self.headers.get("Content-Type", ""))
        if "multipart/form-data" not in content_type:
            self._json(400, {"ok": False, "erro": "Esperado multipart/form-data."})
            return

        fields = _parse_multipart(self)
        arquivo_field = fields.get("arquivo")
        if not arquivo_field or not isinstance(arquivo_field, dict):
            self._json(400, {"ok": False, "erro": "Nenhum arquivo enviado."})
            return

        filename = arquivo_field.get("filename", "")
        file_bytes = arquivo_field.get("data", b"")
        if not filename.lower().endswith(".xlsx"):
            self._json(400, {"ok": False, "erro": "Arquivo invalido. Envie um .xlsx."})
            return

        work_dir = tempfile.mkdtemp()
        input_path = os.path.join(work_dir, "input.xlsx")
        try:
            with open(input_path, "wb") as f:
                f.write(file_bytes)
            from generate_dashboard import load_base, compute_all
            df_raw, plan_name, export_date_str = load_base(input_path)
            df_all, sprint_name, sprint_start, sprint_end, export_date, _sprint_goal = \
                compute_all(df_raw, plan_name, export_date_str)
            total = int(len(df_all))
            if total <= 0:
                self._json(400, {"ok": False, "erro": "Nao foram encontradas tarefas validas."})
                return

            self._json(200, {
                "ok": True,
                "sprint_name": sprint_name or "",
                "total_tarefas": total,
                "sprint_inicio": sprint_start.strftime("%Y-%m-%d") if sprint_start else None,
                "sprint_fim": sprint_end.strftime("%Y-%m-%d") if sprint_end else None,
                "export_date": export_date.strftime("%Y-%m-%d") if export_date else None,
            })
        except Exception as e:
            self._json(400, {"ok": False, "erro": f"Falha ao validar base: {e}"})
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    def _gerar_from_json(self, payload):
        """
        Receives precomputed JSON payload and fills the template.
        """
        template_path = _find_template()
        if not template_path:
            self._json(500, {"erro": "Template nao encontrado."})
            return

        work_dir = tempfile.mkdtemp()
        output_path = os.path.join(work_dir, "dashboard.xlsx")

        try:
            from generate_dashboard import fill_template_from_json

            fill_template_from_json(template_path, payload, output_path)
            if not os.path.exists(output_path):
                self._json(500, {"erro": "Arquivo de saida nao foi gerado."})
                return

            with open(output_path, "rb") as f:
                xlsx_bytes = f.read()

            ts = datetime.now().strftime("%Y%m%d_%H%M")
            out_name = f"Dashboard_{ts}.xlsx"
            warnings = payload.get("warnings", [])

            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Disposition", _safe_filename(out_name))
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self.send_header("X-Dashboard-Warnings", _encode_warnings_header(warnings))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header(
                "Access-Control-Expose-Headers",
                "Content-Disposition, X-Dashboard-Warnings",
            )
            self.end_headers()
            self.wfile.write(xlsx_bytes)

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            msg = str(e)
            if "No sheet named" in msg or "Worksheet" in msg:
                msg = f"Aba nao encontrada no template - {e}"
            self._json(500, {"erro": f"Erro ao preencher template: {msg}"})
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _serve_file(self, path, mime):
        if not os.path.exists(path):
            self._json(404, {"erro": f"Arquivo nao encontrado: {path}"})
            return
        with open(path, "rb") as f:
            content = f.read()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)


def main():
    _setup_shutdown_handler()

    print()
    print("==========================================")
    print(" OnePageReport - Servidor Local ")
    print("==========================================")

    template = _find_template()
    if not template:
        print()
        print("[ERRO] Nenhum template encontrado!")
        print("       Coloque template.next.clean.xlsx (ou template.next.xlsx/template.xlsx)")
        print("       na mesma pasta que server.py.")
        input("\nPressione Enter para sair...")
        return

    print(f"\n[OK] Template : {os.path.basename(template)}")
    print(f"[OK] Porta    : http://localhost:{PORT}")
    print("\n     Abrindo o browser...")
    print("     Para encerrar: feche esta janela ou Ctrl+C\n")

    threading.Timer(1.2, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()

    class ReusableTCPServer(socketserver.TCPServer):
        allow_reuse_address = True

    with ReusableTCPServer(("", PORT), DashboardHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n[OK] Servidor encerrado.")


if __name__ == "__main__":
    main()

"""
OnepageReport Automação - Web Server
Versão compatível com PyInstaller (importa generate_dashboard diretamente).

Uso (desenvolvimento):
    python launcher.py
    Abra http://localhost:5000
"""

import os
import sys
import json
import base64
import shutil
import tempfile
import threading
from datetime import datetime
from flask import Flask, request, send_file, jsonify, after_this_request


# ── Resolve caminhos de recursos (funciona em dev e no .exe compilado) ───────
def _resource(rel):
    """Retorna o caminho absoluto de um recurso bundled ou local."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def _resource_first(*rels):
    """Retorna o primeiro recurso existente na ordem informada."""
    for rel in rels:
        p = _resource(rel)
        if os.path.exists(p):
            return p
    return _resource(rels[0])


# ── Importa o gerador de dashboard ──────────────────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_dashboard import fill_template, compute_json, load_base, compute_all  # noqa: E402

# ── Configuração da aplicação ────────────────────────────────────────────────
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

UPLOAD_DIR    = os.path.join(tempfile.gettempdir(), "onepagereport_work")
TEMPLATE_PATH = _resource_first("template.next.clean.xlsx", "template.next.xlsx", "template.xlsx")
INDEX_HTML    = _resource("index.html")
ALLOWED_EXT   = {".xlsx"}

STARTER_PATH = _resource_first("template.next.clean.xlsx", "template.next.xlsx", "template.xlsx")

REPORT_TYPES = {"equipe", "gp"}

os.makedirs(UPLOAD_DIR, exist_ok=True)


def allowed_file(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT


@app.route("/")
def index():
    with open(INDEX_HTML, encoding="utf-8") as f:
        return f.read()


@app.route("/baixar-starter")
def baixar_starter():
    """Serve o modelo de dados manuais (abas legadas e gp_*)."""
    return send_file(
        STARTER_PATH,
        as_attachment=True,
        download_name="Modelo_Manual.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/validar-base", methods=["POST"])
def validar_base():
    if "arquivo" not in request.files:
        return jsonify({"ok": False, "erro": "Nenhum arquivo enviado."}), 400

    arquivo = request.files["arquivo"]
    if not arquivo.filename:
        return jsonify({"ok": False, "erro": "Nome de arquivo vazio."}), 400
    if not allowed_file(arquivo.filename):
        return jsonify({"ok": False, "erro": "Formato inválido. Envie um .xlsx."}), 400

    work_dir = tempfile.mkdtemp(dir=UPLOAD_DIR)
    input_path = os.path.join(work_dir, "input.xlsx")
    arquivo.save(input_path)

    try:
        df_raw, plan_name, export_date_str = load_base(input_path)
        df_all, sprint_name, sprint_start, sprint_end, export_date, _sprint_goal = \
            compute_all(df_raw, plan_name, export_date_str)
        total = int(len(df_all))
        if total <= 0:
            return jsonify({
                "ok": False,
                "erro": "Arquivo lido, mas não foram encontradas tarefas válidas."
            }), 400

        return jsonify({
            "ok": True,
            "sprint_name": sprint_name or "",
            "total_tarefas": total,
            "sprint_inicio": sprint_start.strftime("%Y-%m-%d") if sprint_start else None,
            "sprint_fim": sprint_end.strftime("%Y-%m-%d") if sprint_end else None,
            "export_date": export_date.strftime("%Y-%m-%d") if export_date else None,
        })
    except Exception as e:
        return jsonify({"ok": False, "erro": f"Falha ao validar base: {e}"}), 400
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/gerar", methods=["POST"])
def gerar():
    if "arquivo" not in request.files:
        return jsonify({"erro": "Nenhum arquivo enviado."}), 400

    arquivo       = request.files["arquivo"]
    author        = request.form.get("autor", "Gerado automaticamente")
    ignore_labels = request.form.get("rotulos_ignorar", "")
    tipo_relatorio = request.form.get("tipo_relatorio", "equipe").strip().lower()
    form_data_raw = request.form.get("dados_formulario", "")
    form_data = None
    if form_data_raw and form_data_raw.strip():
        try:
            form_data = json.loads(form_data_raw)
        except Exception as e:
            return jsonify({"erro": f"dados_formulario inválido: {e}"}), 400

    if not arquivo.filename:
        return jsonify({"erro": "Nome de arquivo vazio."}), 400

    if not allowed_file(arquivo.filename):
        return jsonify({
            "erro": "Formato inválido. Envie um arquivo .xlsx exportado do Microsoft Planner."
        }), 400

    # Tipo de relatório não implementado ainda
    if tipo_relatorio not in REPORT_TYPES:
        return jsonify({"erro": f"Tipo de relatório desconhecido: '{tipo_relatorio}'."}), 400

    if tipo_relatorio == "gp":
        return jsonify({
            "erro": "O Relatório para GP ainda está em desenvolvimento e não está disponível para geração."
        }), 501

    work_dir    = tempfile.mkdtemp(dir=UPLOAD_DIR)
    input_path  = os.path.join(work_dir, "input.xlsx")
    output_path = os.path.join(work_dir, "dashboard.xlsx")
    arquivo.save(input_path)

    # Optional manual data file (3-sheets: Roadmap, Planejamento, Historico)
    dm_path = None
    dm_file = request.files.get("dados_manuais")
    if dm_file and dm_file.filename and allowed_file(dm_file.filename):
        dm_path = os.path.join(work_dir, "dados_manuais.xlsx")
        dm_file.save(dm_path)

    try:
        fill_template(TEMPLATE_PATH, input_path, output_path,
                      author=author,
                      dados_manuais_path=dm_path,
                      report_type=tipo_relatorio,
                      ignore_labels=ignore_labels,
                      form_data=form_data)

        warnings = []
        try:
            data = compute_json(input_path, ignore_labels=ignore_labels)
            warnings = data.get("warnings", [])
        except Exception:
            warnings = []
    except Exception as e:
        shutil.rmtree(work_dir, ignore_errors=True)
        return jsonify({"erro": f"Erro ao gerar dashboard: {e}"}), 500

    if not os.path.exists(output_path):
        shutil.rmtree(work_dir, ignore_errors=True)
        return jsonify({"erro": "O arquivo de saída não foi gerado."}), 500

    @after_this_request
    def cleanup(response):
        try:
            shutil.rmtree(work_dir, ignore_errors=True)
        except Exception:
            pass
        return response

    sprint_name = os.path.splitext(arquivo.filename)[0]
    ts          = datetime.now().strftime("%Y%m%d_%H%M")
    suffix      = "_GP" if tipo_relatorio == "gp" else ""
    out_name    = f"Dashboard{suffix}_{sprint_name}_{ts}.xlsx"

    response = send_file(
        output_path,
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    try:
        encoded_warnings = base64.b64encode(
            json.dumps(warnings, ensure_ascii=False).encode("utf-8")
        ).decode("ascii")
        response.headers["X-Dashboard-Warnings"] = encoded_warnings
        response.headers["Access-Control-Allow-Origin"] = "*"
        response.headers["Access-Control-Expose-Headers"] = (
            "Content-Disposition, X-Dashboard-Warnings"
        )
    except Exception:
        pass
    return response


@app.route("/shutdown", methods=["POST"])
def shutdown():
    """Encerra o servidor Flask (chamado pelo botão 'Fechar' na interface)."""
    def _kill():
        import time
        time.sleep(0.5)          # aguarda a resposta chegar ao browser
        os._exit(0)              # encerra o processo completamente

    threading.Thread(target=_kill, daemon=True).start()
    return jsonify({"ok": True})

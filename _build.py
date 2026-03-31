"""
Script de compilacao do OnepageReport Automacao para .exe
Chamado pelo build.bat - evita problemas de encoding/line endings do batch.
"""
import subprocess
import sys
import os
import shutil

HERE = os.path.dirname(os.path.abspath(__file__))

def main():
    print("\n╔══════════════════════════════════════════════╗")
    print("║     OnepageReport Automacao - Gerador .EXE  ║")
    print("╚══════════════════════════════════════════════╝\n")

    # Verifica arquivos necessarios
    for f in ["launcher.py", "app.py", "generate_dashboard.py", "index.html"]:
        path = os.path.join(HERE, f)
        if not os.path.exists(path):
            print(f"[ERRO] Arquivo necessario nao encontrado: {f}")
            print(f"       Certifique-se de que todos os arquivos estao na mesma pasta.")
            input("\nPressione Enter para sair...")
            sys.exit(1)

    # Determina qual template usar (prefere template.next.xlsx)
    template_next_clean = os.path.join(HERE, "template.next.clean.xlsx")
    template_next = os.path.join(HERE, "template.next.xlsx")
    template_default = os.path.join(HERE, "template.xlsx")
    if os.path.exists(template_next):
        template_file = "template.next.xlsx"
        print(f"[OK] Template: {template_file}")
    elif os.path.exists(template_next_clean):
        template_file = "template.next.clean.xlsx"
        print(f"[OK] Template: {template_file} (fallback clean)")
    elif os.path.exists(template_default):
        template_file = "template.xlsx"
        print(f"[OK] Template: {template_file} (fallback)")
    else:
        print("[ERRO] Nenhum template encontrado (template.next.clean.xlsx, template.next.xlsx ou template.xlsx).")
        input("\nPressione Enter para sair...")
        sys.exit(1)

    # Verifica starter_manual.xlsx (modelo das 3 abas) - opcional
    starter_path = os.path.join(HERE, "starter_manual.xlsx")
    has_starter = os.path.exists(starter_path)
    if has_starter:
        print("[OK] starter_manual.xlsx encontrado - sera incluido no executavel.")
    else:
        print("[info] starter_manual.xlsx nao encontrado - sera ignorado (nao e mais necessario).")

    # Verifica pasta chartex (gráfico termômetro)
    chartex_dir = os.path.join(HERE, "chartex")
    has_chartex = os.path.isdir(chartex_dir) and all(
        os.path.exists(os.path.join(chartex_dir, f))
        for f in ("chartEx1.xml", "chartEx1.xml.rels", "style6.xml", "colors6.xml")
    )
    if has_chartex:
        print("[OK] chartex/ encontrado - grafico termometro sera incluido.")
    else:
        print("[aviso] chartex/ nao encontrado - grafico termometro nao estara disponivel.")

    # Instala dependencias
    print("[1/3] Instalando dependencias...")
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install",
         "pyinstaller", "flask", "openpyxl", "pandas",
         "--quiet", "--disable-pip-version-check"],
        cwd=HERE
    )
    if result.returncode != 0:
        print("[ERRO] Falha ao instalar dependencias. Verifique sua conexao com a internet.")
        input("\nPressione Enter para sair...")
        sys.exit(1)
    print("[OK] Dependencias instaladas.\n")

    # Diretórios de build no disco D (evita fragmentação do SSD do sistema)
    BUILD_DRIVE = "D:\\"
    BUILD_ROOT  = os.path.join(BUILD_DRIVE, "OnepageReport_build")
    DIST_DIR    = os.path.join(BUILD_ROOT, "dist")
    WORK_DIR    = os.path.join(BUILD_ROOT, "work")
    SPEC_DIR    = os.path.join(BUILD_ROOT, "spec")

    # Verifica se o disco D existe; caso contrário usa a pasta local
    if os.path.exists(BUILD_DRIVE):
        print(f"[OK] Disco D encontrado - build sera feito em {BUILD_ROOT}")
        for d in [DIST_DIR, WORK_DIR, SPEC_DIR]:
            os.makedirs(d, exist_ok=True)
    else:
        print("[aviso] Disco D nao encontrado - usando pasta local (dist/).")
        DIST_DIR = os.path.join(HERE, "dist")
        WORK_DIR = os.path.join(HERE, "build")
        SPEC_DIR = HERE

    # Limpa builds anteriores
    exe_anterior = os.path.join(DIST_DIR, "OnepageReport.exe")
    if os.path.isfile(exe_anterior):
        try:
            os.remove(exe_anterior)
        except PermissionError:
            print("[ERRO] O programa OnepageReport.exe ainda esta em execucao.")
            print("       Encerre o programa (botao 'Encerrar programa' no navegador)")
            print("       ou feche a janela do browser e aguarde alguns segundos.")
            input("\nPressione Enter para sair...")
            sys.exit(1)

    for path in [WORK_DIR, os.path.join(SPEC_DIR, "OnepageReport.spec")]:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.isfile(path):
            os.remove(path)

    # Monta argumentos do PyInstaller
    args = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--noconsole",
        "--name", "OnepageReport",
        "--distpath", DIST_DIR,
        "--workpath", WORK_DIR,
        "--specpath", SPEC_DIR,
        "--add-data", f"{os.path.join(HERE, template_file)};.",
        "--add-data", f"{os.path.join(HERE, 'index.html')};.",
        "--add-data", f"{os.path.join(HERE, 'generate_dashboard.py')};.",
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.cell._writer",
        "--hidden-import=openpyxl.styles.stylesheet",
        "--hidden-import=openpyxl.drawing.spreadsheet_drawing",
        "--hidden-import=openpyxl.chart",
        "--hidden-import=pandas",
        "--hidden-import=pandas._libs.tslibs.np_datetime",
        "--hidden-import=pandas._libs.tslibs.nattype",
        "--hidden-import=pandas._libs.tslibs.timedeltas",
        "--hidden-import=flask",
        "--hidden-import=jinja2",
        "--hidden-import=werkzeug",
        "--hidden-import=werkzeug.serving",
        "--hidden-import=werkzeug.routing",
        "--hidden-import=generate_dashboard",
        "--collect-all", "openpyxl",
        "--collect-all", "jinja2",
        os.path.join(HERE, "launcher.py"),
    ]

    # Inclui starter_manual.xlsx se disponível
    if has_starter:
        args += ["--add-data", f"{os.path.join(HERE, 'starter_manual.xlsx')};."]

    # Inclui arquivos chartex se disponíveis
    if has_chartex:
        for fname in ("chartEx1.xml", "chartEx1.xml.rels", "style6.xml", "colors6.xml"):
            args += ["--add-data", f"{os.path.join(HERE, 'chartex', fname)};chartex"]

    print("[2/3] Compilando executavel (pode levar 2-5 minutos)...")
    print("      Aguarde...\n")
    result = subprocess.run(args, cwd=HERE)

    if result.returncode != 0:
        print("\n[ERRO] Falha na compilacao.")
        input("\nPressione Enter para sair...")
        sys.exit(1)

    # Limpa arquivos temporarios
    print("\n[3/3] Limpando arquivos temporarios...")
    for path in [WORK_DIR, os.path.join(SPEC_DIR, "OnepageReport.spec")]:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.isfile(path):
            os.remove(path)

    exe_path = os.path.join(DIST_DIR, "OnepageReport.exe")
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        exe_short = exe_path.replace(os.path.expanduser("~"), "~")
        print(f"\n╔══════════════════════════════════════════════════════╗")
        print(f"║   PRONTO! Executavel gerado com sucesso.             ║")
        print(f"║                                                      ║")
        print(f"║   Arquivo: {os.path.basename(exe_path)} ({size_mb:.0f} MB){'':<18}║")
        print(f"║   Pasta  : {DIST_DIR[:44]:<44} ║")
        print(f"║                                                      ║")
        print(f"║   Copie o .exe para qualquer pasta e                 ║")
        print(f"║   duplo clique para usar. Sem instalar nada!         ║")
        print(f"╚══════════════════════════════════════════════════════╝\n")

        resp = input("Deseja abrir a pasta com o executavel? (S/N): ").strip().lower()
        if resp == "s":
            subprocess.Popen(["explorer", DIST_DIR])
    else:
        print(f"[AVISO] Executavel nao encontrado em: {exe_path}")

    input("\nPressione Enter para sair...")

if __name__ == "__main__":
    main()

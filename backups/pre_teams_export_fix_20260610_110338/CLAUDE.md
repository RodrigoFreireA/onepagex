# Contexto tecnico do projeto

## Marco estavel atual

Versao de referencia consolidada em **31/03/2026 10:25 (America/Sao_Paulo)**.

Backups desta versao:
- `backups/stable_complete_20260331_102539/`
- `backups/stable_complete_20260331_102539.zip`

Orientacao: antes de alterar parsing/preenchimento de Excel ou fluxo do formulario, criar novo snapshot em `backups/` e testar os 3 modos (`none`, `dm`, `form`).

Este projeto gera um Excel final de dashboard unindo:
- dados brutos exportados do Planner/Teams
- dados manuais `gp_*` (planilha manual ou formulario)

Objetivo: manter graficos e formulas do template, mudando apenas dados de entrada.

## Arquitetura atual

- `index.html`
  - wizard com 3 modos (`none`, `dm`, `form`)
  - validacao da base antes de liberar etapa manual em `dm` e `form`
  - salva ultimo formulario em `localStorage` (`onepagereport.form.last.v1`)
- `app.py`
  - `POST /validar-base`: valida planner export e retorna resumo
  - `POST /gerar`: gera `Dashboard_*.xlsx`
  - `GET /baixar-starter`: baixa modelo manual
- `generate_dashboard.py`
  - parse da base planner
  - calculo das tabelas/KPIs
  - injecao de planilha manual (`dados_manuais`)
  - injecao de formulario (`dados_formulario`)
  - pos-processamento ZIP para preservar graficos

## Contrato de fluxo (nao quebrar)

1. Escolher modo na tela inicial.
2. Upload da base planner.
3. Validar base (`/validar-base`) se modo for `dm` ou `form`.
4. Coletar dados manuais:
   - `dm`: upload do excel manual.
   - `form`: montar JSON do formulario.
5. Gerar (`/gerar`) e baixar arquivo final.

## Principais mudancas ja aplicadas

- Fluxo visual separado por modo (pagina inicial de escolha + workflow).
- Validacao de base obrigatoria para modos com dados manuais.
- Suporte a todas as abas `gp_*` no upload manual e no formulario.
- Filtro de rotulos ignorados (`rotulos_ignorar`).
- Correcao de corrupcao de arquivo no Excel com restauracao de partes de grafico no ZIP.
- Limpeza de referencias especificas de projeto antigo no front para uso generico.

## Regras de seguranca para nao corromper o XLSX

- Sempre executar `_postprocess_xlsx(output_path, template_path=template_path)`.
- Nao remover copia de `xl/charts/*` e dependencias (`style*.xml`, `colors*.xml`) quando o template as referencia.
- Nao alterar manualmente `drawing*.xml` e `.rels` sem atualizar dependencias.
- Nao mudar nomes de abas esperadas (`x...`, `gp_...`) sem atualizar o codigo.
- Nao substituir `template.next.clean.xlsx` por template incompleto.

## Checklist antes de entregar alteracao

- Gerar 1 arquivo no modo `none`.
- Gerar 1 arquivo no modo `dm` com planilha manual.
- Gerar 1 arquivo no modo `form`.
- Abrir no Excel e confirmar:
  - sem aviso de reparo/corrupcao
  - graficos visiveis
  - abas `x*` preenchidas
  - abas `gp_*` refletindo entrada manual

## Comandos uteis

```bash
pip install -r requirements.txt
python launcher.py
```

```bash
python -m py_compile app.py server.py generate_dashboard.py
```

## Limites atuais conhecidos

- `tipo_relatorio=gp` ainda retorna `501` (nao implementado para geracao final).

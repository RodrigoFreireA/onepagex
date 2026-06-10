# OnePageReport Automacao

Projeto para gerar um dashboard Excel a partir de:
- base bruta exportada do Planner/Teams (`.xlsx`)
- dados manuais opcionais (`gp_*`) via planilha manual ou formulario web

## Estado atual (o que foi feito)

- Fluxo visual em etapas com tela inicial de escolha de modo:
  - `Sem dados manuais`
  - `Com planilha manual`
  - `Com formulario`
- Validacao da planilha base antes da geracao para modos `Com planilha manual` e `Com formulario` (`POST /validar-base`).
- Suporte completo a abas manuais `gp_*`:
  - `gp_Plann_Projeto`
  - `gp_Plann_Sprint`
  - `gp_Cabecalhos`
  - `gp_Disponibilidade`
  - `gp_Transversalidade`
  - `gp_Roadmap`
- Persistencia do ultimo formulario no navegador (`localStorage`) para pre-preenchimento automatico na proxima execucao.
- Filtro de rotulos ignorados (`rotulos_ignorar`) no processamento.
- Correcao da corrupcao pos-geracao do Excel:
  - pos-processamento ZIP preservando dependencias de graficos (`chart`, `chartEx`, `style`, `colors`)
  - preservacao de visuais/transparencia dos graficos do template

## Fluxo completo de uso

1. Usuario escolhe o modo na primeira tela.
2. Usuario envia o arquivo base exportado do Planner (`.xlsx`).
3. Se modo for `Com planilha manual` ou `Com formulario`, o sistema valida a base.
4. Usuario completa a parte manual:
   - `Com planilha manual`: envia segunda planilha com abas `gp_*`.
   - `Com formulario`: preenche campos/tabelas no front.
5. Usuario clica em `Gerar Dashboard`.
6. Backend processa os dados e preenche o template.
7. Sistema baixa o arquivo final `Dashboard_<nome_sprint>_<timestamp>.xlsx`.

## O que evitar para nao bugar novamente

- Nao remover a chamada de `_postprocess_xlsx(...)` ao final da geracao.
- Nao trocar o template principal por arquivo sem os assets de grafico.
- Nao renomear abas esperadas no template (`x*` e `gp_*`), senao os preenchimentos quebram.
- Nao salvar manualmente o template abrindo/fechando em ferramentas que removem partes XML de grafico.
- Nao alterar os nomes de campos enviados pelo frontend (`dados_manuais`, `dados_formulario`, `rotulos_ignorar`, `arquivo`).
- Nao ignorar a validacao da base em modos que exigem mesclagem manual.

## Execucao local

```bash
pip install -r requirements.txt
python launcher.py
```

Alternativa:

```bash
python server.py
```

## Estrutura minima da pasta

- `generate_dashboard.py` - logica principal de extracao/calculo/preenchimento
- `app.py` - backend Flask (fluxo principal)
- `server.py` - servidor HTTP alternativo
- `index.html` - frontend (wizard + formulario)
- `template.next.clean.xlsx` - template base limpo e pronto para preencher

## Troubleshooting rapido

- Erro `Encontramos um problema em um conteudo...` no Excel:
  - validar se `template.next.clean.xlsx` esta presente e correto
  - confirmar que o pos-processamento ZIP nao foi removido
  - regenerar saida e testar em arquivo novo
- Grafico nao aparece:
  - conferir se a aba de origem (`x*`) recebeu dados
  - conferir se os ranges/nomes do template nao foram alterados manualmente

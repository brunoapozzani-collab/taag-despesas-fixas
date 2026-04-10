# TAAG — Despesas Fixas

Web app para gerar o resumo executivo de despesas fixas das empresas TAAG.

## Como rodar localmente

```bash
cd Thais.2
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/streamlit run tools/app.py
```

Abra o navegador em `http://localhost:8501`.

## Como usar

1. **Upload** da planilha mensal na barra lateral.
2. Escolha o **período** (data inicial e final).
3. Aba **Revisar Despesas**: confira a classificação automática. Marque/desmarque cada linha como "Fixa". Use os filtros por empresa/categoria.
4. Aba **Resumo**: prévia dos totais por empresa.
5. Aba **Gerar Relatórios**: clique em **Confirmar e Gerar** e baixe o PDF executivo + a planilha detalhada.

O preset (lista de palavras-chave + ajustes manuais) é salvo automaticamente para a próxima sessão.

## Deploy no Streamlit Community Cloud

1. Empurre este repositório para o GitHub.
2. Em [share.streamlit.io](https://share.streamlit.io), conecte o repositório.
3. Aponte para `tools/app.py` como entrada.
4. O Streamlit Cloud instala `requirements.txt` automaticamente.

## Estrutura

```
Thais.2/
├── tools/
│   ├── app.py             # Streamlit UI
│   ├── expense_engine.py  # Carregamento, limpeza, classificação
│   ├── pdf_report.py      # Gerador do PDF executivo
│   └── excel_report.py    # Gerador do Excel detalhado
├── workflows/
│   └── fixed_expenses_report.md
├── data/
│   ├── logo.png           # Logo TAAG
│   ├── symbol.png
│   ├── plano_de_contas.xlsx
│   └── presets.json       # criado na primeira execução
├── requirements.txt
└── .streamlit/config.toml
```

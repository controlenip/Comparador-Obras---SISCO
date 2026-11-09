# NIP | Comparador de Levantamentos

Aplicação Streamlit com layout inspirado no Dashboard NIP: faixa de cabeçalho azul-marinho com logomarca, filtros operacionais à esquerda e navegação horizontal por abas.

## Arquivo principal

`app.py` — definir como **Main file path** no Streamlit Community Cloud.

## Publicação

1. Criar ou abrir o repositório no GitHub.
2. Enviar `app.py`, `comparador.py`, `requirements.txt`, `.streamlit/config.toml` e a pasta `assets/` (manter a estrutura).
3. No Streamlit Community Cloud, selecionar repositório e branch `main`, apontando para `app.py`.
4. Carregar a planilha anterior e a atual nos filtros à esquerda.
5. Consultar Dashboard, Entrantes por dia, Obras entrantes, Auditoria e Exportação.

## Regras da comparação

- Coluna **A**: ID do levantamento, identificador principal.
- Coluna **F**: Nota CCS, conferência de continuidade.
- Coluna **O**: tipo de projeto, conferência de continuidade.
- Coluna **AA**: data do recebimento, filtro de ano (padrão 2026).
- Uma obra é entrante quando é encontrada na base atual mas não na anterior, e sua data AA é do ano escolhido.
- Se faltar ID em A, usa-se F+O como chave de contingência. Quando F+O é único nas duas bases, também identifica mudanças de ID sem tratá-las como entrantes.
- Datas inválidas, duplicidades e novas obras fora do ano são apresentadas na auditoria.
- Projeção mensal estimada a partir dos entrantes observados numa janela de até 30 dias (inclusive dias sem registros). É uma extrapolação, não uma previsão estatística robusta nem o histórico total do mês.

## Testes

`python -m pip install -r requirements.txt pytest`

`python -m pytest -q tests/test_comparador.py`

## Privacidade

Os arquivos enviados são processados pela sessão Streamlit em execução; não há integração de envio de conteúdo a APIs externas no código do comparador. Evite colocar planilhas com dados de clientes no repositório público.

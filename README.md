# NIP | Comparador de Levantamentos

Aplicacao Streamlit para comparar **dois snapshots** da planilha `Levantamentos`, identificar obras novas com data de recebimento em **2026** (ou outro ano selecionado), distinguir alteracoes e exportar Excel com as linhas completas.

## Como publicar no GitHub e Streamlit Cloud
1. Crie um repositorio privado (recomendado para dados operacionais).
2. Envie todos os arquivos **incluindo `app.py`, `comparador.py`, `requirements.txt`, `.streamlit/config.toml` e `assets/logo_nip.png`**.
3. Em [share.streamlit.io](https://share.streamlit.io/) clique em **Create app** e selecione o repositorio, branch `main` e arquivo principal `app.py`.
4. Execute e envie **primeiro a planilha anterior** e **depois a planilha atual**. Escolha o ano (padrao 2026).
5. Exporte o Excel gerado. Os arquivos fornecidos ao aplicativo nao sao salvos no repositorio.

### Regras
- Coluna **A** (`ID do levantamento`) e identificador principal. Se A estiver vazia, chave secundaria **F + O** (`Nota CCS` + `Tipo de Projeto(PI)`). Se A e F estiverem vazias, a linha e sinalizada.
- A comparacao usa a **base anterior completa** como referencia, mesmo quando AA nela for de outro ano. Isso evita falsos entrantes.
- So e considerado **entrante** um registro que nao existe na base anterior, existe na atual e cuja **AA** (`Recebimento(data criacao)`) esteja no ano escolhido.
- Alteracoes nos campos A/F/O/AA em registros ja identificados nao contam como novas obras. Quando o ID A muda, nao ha como identificar com certeza que seja a mesma obra; revise as inconsistencias.
- Duplicatas pelo identificador sao contabilizadas uma so vez e sinalizadas.
- O Excel exporta todas as colunas originais dos entrantes e abas de auditoria.
- A media diaria e calculada a partir das **datas AA dos entrantes identificados**, incluindo dias sem entradas dentro da janela observada. A projecao padrao representa **30 dias** e e indicativa, nao uma previsao estatistica robusta nem o volume total de obras do mes. Um intervalo de snapshots muito curto pode distorcer a estimativa.
- A AA e a data informada no SISCO/planilha, **nao** necessariamente a data em que a linha passou a aparecer entre os dois arquivos.

### Uso local
```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

### Seguranca
Nao coloque planilhas reais no repositorio. A analise ocorre na sessao Streamlit; se hospedar em Streamlit Community Cloud, os uploads passam pela infraestrutura do provedor. Configure controle de acesso se forem dados confidenciais.

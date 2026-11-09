"""NIP | Comparador de Levantamentos - dashboard Streamlit."""
from __future__ import annotations
import base64
import html
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from comparador import analyze, daily_counts, monthly_forecast, make_excel, dt, COL_A, COL_F, COL_O, COL_AA

st.set_page_config(
    page_title='NIP | Comparador de Levantamentos',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='collapsed'
)
ROOT = Path(__file__).resolve().parent
logo = base64.b64encode((ROOT / 'assets/logo_nip.png').read_bytes()).decode('ascii')

st.markdown('''<style>
:root{--navy:#143b54;--teal:#0b7785;--teal2:#14889c;--ink:#15384f;--line:#d5e3ed;--muted:#6c8492;--bg:#edf2f6}
html,body,[data-testid="stAppViewContainer"]{background:var(--bg)!important}
header[data-testid="stHeader"]{height:0!important;min-height:0!important;background:transparent!important}
div[data-testid="stToolbar"], [data-testid="stDecoration"], #MainMenu, footer{display:none!important}
[data-testid="stSidebar"],[data-testid="collapsedControl"],[data-testid="stSidebarCollapsedControl"]{display:none!important}
[data-testid="stMainBlockContainer"],.block-container{width:100%!important;max-width:none!important;padding:0 1.7rem 2rem!important}
[data-testid="stMainBlockContainer"]{padding-top:0!important}
[data-testid="stVerticalBlock"]{gap:.7rem}

.nip-head{background:var(--navy);display:flex;align-items:center;gap:18px;padding:10px 22px 10px 20px;margin:0 -1.7rem 10px;min-height:88px;color:#fff;overflow:visible}
.nip-head img{width:94px;height:66px;object-fit:contain;display:block;flex-shrink:0}
.nip-head-main{display:flex;flex-direction:column;justify-content:center;min-width:0}
.nip-head-title{font-size:18px;font-weight:800;color:#fff;letter-spacing:-.15px;line-height:1.12}
.nip-head-sub{font-size:11px;color:#d8e8ef;margin-top:7px;line-height:1.25}
.nip-head-right{margin-left:auto;background:#1393a2;padding:10px 16px;border-radius:5px;font-size:12px;font-weight:800;white-space:nowrap;box-shadow:0 2px 8px rgba(0,0,0,.12);margin-right:4px}

.nip-rail{font-size:14px;font-weight:800;color:var(--ink);margin:6px 0 12px}
.nip-sub{font-size:11px;color:var(--muted);margin:6px 0 10px;line-height:1.35}
.nip-rail-foot{border-top:1px solid #d9e5ec;padding-top:16px;margin-top:18px;color:#7b909c;font-size:11px;line-height:1.4}

/* Navegacao fiel ao dashboard NIP: a selecao e gerenciada pelo st.tabs. */
[data-testid="stTabs"]{background:transparent!important;min-width:0}
[data-testid="stTabs"] [role="tablist"]{
    display:flex!important;
    align-items:center!important;
    flex-wrap:wrap!important;
    gap:8px!important;
    height:auto!important;
    min-height:53px!important;
    margin:0 0 13px!important;
    padding:9px 11px!important;
    background:#f6f9fc!important;
    border:1px solid #c9dce8!important;
    border-radius:7px!important;
    box-shadow:0 1px 2px rgba(16,48,77,.035)!important;
}
[data-testid="stTabs"] [role="tablist"] button[role="tab"]{
    display:inline-flex!important;
    align-items:center!important;
    justify-content:center!important;
    flex:0 0 auto!important;
    box-sizing:border-box!important;
    min-width:0!important;
    min-height:34px!important;
    height:34px!important;
    margin:0!important;
    padding:0 14px!important;
    border:1px solid #a5c4d6!important;
    border-radius:7px!important;
    background:#fff!important;
    color:#17365d!important;
    font-size:12px!important;
    font-weight:700!important;
    line-height:1!important;
    white-space:nowrap!important;
    box-shadow:none!important;
    transition:background-color .14s ease,border-color .14s ease,color .14s ease,box-shadow .14s ease!important;
    transform:none!important;
    text-decoration:none!important;
}
[data-testid="stTabs"] [role="tablist"] button[role="tab"] :is(p,span){
    color:inherit!important;
    font-size:inherit!important;
    font-weight:inherit!important;
    line-height:inherit!important;
    margin:0!important;
}
[data-testid="stTabs"] [role="tablist"] button[role="tab"]:not([aria-selected="true"]):hover{
    background:#eaf2f7!important;
    border-color:#7da6bd!important;
    color:#123b56!important;
}
[data-testid="stTabs"] [role="tablist"] button[role="tab"][aria-selected="true"],
[data-testid="stTabs"] [role="tablist"] button[role="tab"][aria-selected="true"]:hover,
[data-testid="stTabs"] [role="tablist"] button[role="tab"][aria-selected="true"]:focus{
    background:#143b54!important;
    border-color:#143b54!important;
    color:#fff!important;
    box-shadow:0 2px 5px rgba(20,59,84,.18)!important;
    transform:none!important;
}
[data-testid="stTabs"] [role="tablist"] button[role="tab"]:focus-visible{
    outline:2px solid #56afc8!important;
    outline-offset:2px!important;
}
[data-testid="stTabs"] [data-baseweb="tab-highlight"],
[data-testid="stTabs"] [data-baseweb="tab-border"]{display:none!important}

[data-testid="stFileUploader"]{border:1px solid #d5e3ed!important;border-radius:9px;background:#fff;padding:6px}
[data-testid="stFileUploaderDropzone"]{padding:9px!important;min-height:68px!important}
[data-testid="stFileUploaderDropzoneInstructions"]{font-size:11px!important}
[data-testid="stFileUploader"] section button{font-size:11px!important}

.nip-kpi{background:#fff;border:1px solid #d5e3ed;border-radius:10px;padding:14px 17px 15px;min-height:118px;position:relative;overflow:hidden;box-shadow:0 1px 1px rgba(12,44,66,.02)}
.nip-kpi:before{content:'';position:absolute;left:0;right:0;bottom:0;height:3px;background:#0d7785}
.nip-kpi.blue:before{background:#3c82f6}.nip-kpi.green:before{background:#47bd86}.nip-kpi.gold:before{background:#e3a527}.nip-kpi.teal:before{background:#0d7785}.nip-kpi.slate:before{background:#5f89a4}
.nip-kpi-label{font-size:11px;color:#6f8794;font-weight:600;line-height:1.25}
.nip-kpi-value{font-size:21px;color:#123c59;font-weight:800;line-height:1.05;margin-top:14px;letter-spacing:-.2px}
.nip-kpi-note{font-size:11px;color:#6f8794;line-height:1.3;margin-top:13px}
.nip-kpi-highlight{background:linear-gradient(180deg,#f7fbfd 0%,#ffffff 100%);border-color:#cae0ea}

.nip-panel{background:#fff;border:1px solid #d5e3ed;border-radius:10px;padding:14px 16px;margin:8px 0 12px}
.nip-panel-title{color:var(--ink);font-weight:800;font-size:14px;margin:1px 0 12px}
.nip-panel-sub{color:#7a919d;font-size:11px;line-height:1.35}
.nip-intro{color:#5e7886;font-size:12px;margin:1px 0 7px}
.nip-note{font-size:11px;color:#7a919d;line-height:1.45;margin-top:9px}

.nip-mini{background:#fff;border:1px solid #d9e6ed;padding:14px 15px;border-radius:9px;min-height:100px}
.nip-mini .t{font-size:11px;color:#6f8794;font-weight:600}
.nip-mini .v{font-size:20px;color:#123c59;font-weight:800;margin-top:10px;line-height:1.05}
.nip-mini .s{font-size:11px;color:#758d99;margin-top:8px;line-height:1.3}

[data-testid="stDataFrame"],div[data-testid="stPlotlyChart"]{border:1px solid #d7e4ed;border-radius:9px;background:#fff;padding:5px}
[data-testid="stMetric"]{border:1px solid #d9e6ed;padding:12px;background:#fff;border-radius:8px}
[data-testid="stAlert"]{border-radius:8px}
[data-testid="stVerticalBlockBorderWrapper"]{border-radius:8px}

@media(max-width:1100px){
 [data-testid="stMainBlockContainer"],.block-container{padding:0 12px 20px!important}
 .nip-head{margin:0 -12px 10px;padding:10px 14px;gap:12px}
 .nip-head img{width:74px;height:50px}
 .nip-head-right{display:none}
 .nip-head-title{font-size:16px}
}
</style>''', unsafe_allow_html=True)

st.markdown(
    f'''<div class="nip-head">
    <img src="data:image/png;base64,{logo}" alt="NIP Grupo Ígneo">
    <div class="nip-head-main">
        <div class="nip-head-title">NIP | Comparador de Levantamentos</div>
        <div class="nip-head-sub">Obras entrantes · Comparação de bases · Análise diária · Projeção mensal</div>
    </div>
    <div class="nip-head-right">✓ Comparação local de arquivos Excel</div>
    </div>''',
    unsafe_allow_html=True
)

@st.cache_data(show_spinner='Comparando as planilhas...', max_entries=2)
def cached_analysis(old: bytes, new: bytes, year: int):
    return analyze(old, new, year=year)

@st.cache_data(show_spinner='Preparando arquivo...', max_entries=2)
def cached_excel(result: dict):
    return make_excel(result)

left, right = st.columns([1.07, 4.93], gap='medium', vertical_alignment='top')
with left:
    st.markdown('<div class="nip-rail">Filtros operacionais</div>', unsafe_allow_html=True)
    previous = st.file_uploader('1 · Levantamentos anterior', type=['xlsx'], key='previous', help='Arquivo anterior para comparação.')
    current = st.file_uploader('2 · Levantamentos atual', type=['xlsx'], key='current', help='Arquivo mais recente.')
    year = st.number_input('Ano do recebimento', min_value=2020, max_value=2100, value=2026, step=1)
    st.markdown('<div class="nip-sub">Selecione as duas versões da planilha para identificar somente as obras novas.</div>', unsafe_allow_html=True)
    st.markdown('<div class="nip-rail-foot">NIP · Grupo Ígneo</div>', unsafe_allow_html=True)

with right:
    tabs = st.tabs(['Dashboard', 'Entrantes por dia', 'Obras entrantes', 'Auditoria', 'Exportação'])
    if not previous or not current:
        with tabs[0]:
            st.markdown('<div class="nip-panel-title">◉ Resumo da comparação</div>', unsafe_allow_html=True)
            st.info('Envie as duas versões da planilha Levantamentos nos filtros à esquerda para iniciar a comparação.')
            x, y, z = st.columns(3)
            cards = [
                (x, '01 · Comparar', 'Localizar registros presentes apenas na base atual.'),
                (y, '02 · Analisar', 'Contabilizar o volume de novas entradas por dia e por período.'),
                (z, '03 · Exportar', 'Gerar uma terceira planilha somente com as obras entrantes.'),
            ]
            for col, title, desc in cards:
                with col:
                    st.markdown(
                        f'<div class="nip-kpi teal nip-kpi-highlight"><div class="nip-kpi-label">{title}</div><div class="nip-kpi-value" style="font-size:15px;line-height:1.35">{desc}</div></div>',
                        unsafe_allow_html=True,
                    )
        with tabs[1]:
            st.info('Envie os arquivos para visualizar a evolução diária das entradas.')
        with tabs[2]:
            st.info('Envie os arquivos para pesquisar as novas obras identificadas.')
        with tabs[3]:
            st.info('Envie os arquivos para verificar auditoria, alterações e duplicidades.')
        with tabs[4]:
            st.info('Envie os arquivos para gerar a planilha de exportação.')
        st.stop()

    if previous.getvalue() == current.getvalue():
        st.warning('Os dois arquivos enviados são idênticos. Nenhuma diferença será identificada.')
    try:
        result = cached_analysis(previous.getvalue(), current.getvalue(), int(year))
    except Exception as exc:
        st.error(f'Erro ao analisar os arquivos: {exc}')
        st.stop()

    entrants = result['entrants']
    counts = daily_counts(entrants)
    forecast = monthly_forecast(entrants)

    def frame_rows(rows, with_reason=False):
        headers = result['headers']
        return pd.DataFrame(
            ([num, reason, *row] if with_reason else [num, *row] for num, row, reason in rows),
            columns=['Linha original'] + (['Detalhe'] if with_reason else []) + headers,
        )

    def kpi(label, value, sub='', variant='teal'):
        return (
            f'<div class="nip-kpi {variant}">'
            f'<div class="nip-kpi-label">{html.escape(label)}</div>'
            f'<div class="nip-kpi-value">{value}</div>'
            f'<div class="nip-kpi-note">{html.escape(sub)}</div>'
            f'</div>'
        )

    type_series = pd.Series(
        [str(row[COL_O]).strip() if row[COL_O] not in (None, '') else 'Não informado' for _, row, _ in entrants]
    ).value_counts() if entrants else pd.Series(dtype='int64')
    top_type = type_series.index[0] if len(type_series) else 'Sem registros'
    top_type_count = int(type_series.iloc[0]) if len(type_series) else 0

    with tabs[0]:
        st.markdown('<div class="nip-panel-title">◉ Resumo de entrantes</div>', unsafe_allow_html=True)
        row1 = st.columns(5, gap='small')
        info1 = [
            ('Base anterior', f"{result['old_count']:,}".replace(',', '.'), 'Registros analisados', 'blue'),
            ('Base atual', f"{result['new_count']:,}".replace(',', '.'), 'Registros analisados', 'blue'),
            ('Novos entrantes', f"{len(entrants):,}".replace(',', '.'), 'Obras novas identificadas', 'green'),
            ('Obras alteradas', f"{len(result['changes']):,}".replace(',', '.'), 'Mudanças em registros já existentes', 'teal'),
            ('Projeção mensal', f"{forecast['projection']:,}".replace(',', '.'), 'Estimativa indicativa para 30 dias', 'teal'),
        ]
        for col, (label, value, sub, variant) in zip(row1, info1):
            with col:
                st.markdown(kpi(label, value, sub, variant), unsafe_allow_html=True)

        row2 = st.columns([1.15, 1, 1], gap='small')
        with row2[0]:
            st.markdown(
                kpi('Tipo de obra com maior entrada', html.escape(top_type), f'{top_type_count:,} obra(s) no período'.replace(',', '.'), 'gold'),
                unsafe_allow_html=True,
            )
        with row2[1]:
            avg_txt = f"{forecast['avg']:.2f}".replace('.', ',')
            st.markdown(kpi('Média diária', avg_txt, 'Média observada no período', 'slate'), unsafe_allow_html=True)
        with row2[2]:
            st.markdown(kpi('Dias observados', str(forecast['days']), 'Janela usada no cálculo', 'slate'), unsafe_allow_html=True)

        chartcol, statscol = st.columns([1.55, 1], gap='small')
        with chartcol:
            st.markdown('<div class="nip-panel-title">▥ Evolução diária dos entrantes</div>', unsafe_allow_html=True)
            if counts:
                days = pd.date_range(min(counts), max(counts), freq='D')
                daily = pd.DataFrame({'Data': days, 'Entrantes': [counts.get(day.date(), 0) for day in days]})
                fig = px.bar(daily, x='Data', y='Entrantes', color_discrete_sequence=['#0d7785'])
                fig.update_layout(
                    plot_bgcolor='white', paper_bgcolor='white', font_color='#143b54',
                    margin=dict(l=8, r=8, t=12, b=5), height=320,
                    xaxis_title='', yaxis_title='',
                )
                fig.update_traces(hovertemplate='Data: %{x|%d/%m/%Y}<br>Entrantes: %{y}<extra></extra>')
                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
            else:
                st.info('Não houve obras entrantes no período filtrado.')
        with statscol:
            st.markdown('<div class="nip-panel-title">◉ Indicadores de tendência</div>', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(
                    f'<div class="nip-mini"><div class="t">Média diária</div><div class="v">{forecast["avg"]:.2f}'.replace('.', ',') + '</div><div class="s">Entradas por dia</div></div>',
                    unsafe_allow_html=True,
                )
            with c2:
                st.markdown(
                    f'<div class="nip-mini"><div class="t">Dias observados</div><div class="v">{forecast["days"]}</div><div class="s">Período considerado</div></div>',
                    unsafe_allow_html=True,
                )
            st.markdown(
                f'<div class="nip-mini" style="margin-top:10px"><div class="t">Entrantes na janela</div><div class="v">{forecast["total"]:,}</div><div class="s">Volume apurado no período observado</div></div>'.replace(',', '.'),
                unsafe_allow_html=True,
            )
            if forecast['period']:
                start, end = forecast['period']
                st.markdown(f'<div class="nip-note">Período observado: {start:%d/%m/%Y} a {end:%d/%m/%Y}</div>', unsafe_allow_html=True)
            st.markdown('<div class="nip-note">Projeção indicativa de 30 dias com base no comportamento observado entre os arquivos comparados.</div>', unsafe_allow_html=True)

        st.markdown('<div class="nip-panel-title">▥ Distribuição por tipo de obra</div>', unsafe_allow_html=True)
        if entrants:
            serie = type_series.head(15)
            fig = px.bar(
                x=serie.values, y=serie.index, orientation='h',
                labels={'x': 'Entrantes', 'y': 'Tipo de obra'}, color_discrete_sequence=['#0d7785']
            )
            fig.update_layout(
                yaxis={'autorange': 'reversed'}, height=max(240, len(serie) * 30 + 65),
                plot_bgcolor='white', paper_bgcolor='white', margin=dict(l=10, r=20, t=5, b=5),
                xaxis_title='', yaxis_title=''
            )
            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        else:
            st.info('Sem obras novas para detalhar a distribuição por tipo.')

    with tabs[1]:
        st.markdown('<div class="nip-panel-title">▥ Entradas diárias e consolidação mensal</div>', unsafe_allow_html=True)
        if counts:
            diary = pd.DataFrame([{'Data': day, 'Entrantes': qty} for day, qty in sorted(counts.items())])
            diary['Mês'] = diary['Data'].apply(lambda d: d.strftime('%m/%Y'))
            months = sorted(diary['Mês'].unique(), key=lambda v: (int(v[3:]), int(v[:2])), reverse=True)
            chosen = st.selectbox('Mês para detalhamento', months)
            filtered = diary[diary['Mês'] == chosen]
            st.markdown(kpi('Entrantes no mês selecionado', int(filtered['Entrantes'].sum()), 'Total no mês escolhido', 'green'), unsafe_allow_html=True)
            st.dataframe(filtered[['Data', 'Entrantes']].rename(columns={'Data': 'Data', 'Entrantes': 'Quantidade'}), hide_index=True, use_container_width=True, height=340)
            st.markdown('<div class="nip-panel-title">Consolidado mensal</div>', unsafe_allow_html=True)
            st.dataframe(diary.groupby('Mês', as_index=False)['Entrantes'].sum().rename(columns={'Entrantes': 'Quantidade'}), hide_index=True, use_container_width=True)
        else:
            st.info('Sem entrantes para o ano escolhido.')

    with tabs[2]:
        st.markdown('<div class="nip-panel-title">▥ Relação de novas obras</div>', unsafe_allow_html=True)
        term = st.text_input('Pesquisar obra, nota ou tipo', placeholder='Número da obra, CCS ou tipo...')
        dates = sorted({dt(row[COL_AA]) for _, row, _ in entrants if dt(row[COL_AA])})
        selected = st.multiselect('Filtrar datas', dates, format_func=lambda day: day.strftime('%d/%m/%Y'))
        filtered = []
        for item in entrants:
            _, row, _ = item
            if selected and dt(row[COL_AA]) not in selected:
                continue
            if term and not any(term.casefold() in str(row[i] or '').casefold() for i in (COL_A, COL_F, COL_O)):
                continue
            filtered.append(item)
        st.caption(f'{len(filtered):,} obras exibidas · {len(entrants):,} entrantes identificados'.replace(',', '.'))
        if filtered:
            frame = frame_rows(filtered)
            names = result['headers']
            st.dataframe(frame[['Linha original', names[COL_A], names[COL_F], names[COL_O], names[COL_AA]]].head(10000), hide_index=True, use_container_width=True, height=470)
            with st.expander('Prévia com todas as colunas originais'):
                st.dataframe(frame.head(200), hide_index=True, use_container_width=True)
            st.download_button(
                'Baixar entrantes filtrados (.xlsx)',
                make_excel(result, filtered),
                'NIP_ENTRANTES_FILTRADOS.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
        else:
            st.info('Nenhuma obra corresponde aos filtros aplicados.')

    with tabs[3]:
        st.markdown('<div class="nip-panel-title">◉ Auditoria da comparação</div>', unsafe_allow_html=True)
        columns = st.columns(3)
        metrics = [
            ('Duplicados atuais', len(result['duplicates'])),
            ('Novos fora do período', len(result['excluded'])),
            ('Sem identificação / data', len(result['invalid'])),
        ]
        for col, (label, value) in zip(columns, metrics):
            with col:
                st.markdown(kpi(label, f'{value:,}'.replace(',', '.'), 'Registros para revisão', 'teal'), unsafe_allow_html=True)
        st.caption('Alterações identificadas em registros já existentes são tratadas separadamente e não entram como novas obras.')
        with st.expander('Registros existentes alterados'):
            st.dataframe(frame_rows(result['changes']).head(500) if result['changes'] else pd.DataFrame(), hide_index=True, use_container_width=True)
        with st.expander('Registros fora do período, duplicidades e inconsistências'):
            records = result['excluded'] + result['invalid'] + result['duplicates']
            st.dataframe(frame_rows(records, with_reason=True).head(500) if records else pd.DataFrame(), hide_index=True, use_container_width=True)

    with tabs[4]:
        st.markdown('<div class="nip-panel-title">▤ Exportação da terceira planilha</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="nip-kpi green"><div class="nip-kpi-label">Obras novas localizadas</div><div class="nip-kpi-value">{len(entrants):,}</div><div class="nip-kpi-note">Arquivo Excel com os entrantes e a auditoria da comparação</div></div>'.replace(',', '.'),
            unsafe_allow_html=True,
        )
        st.download_button(
            '⬇ Baixar planilha completa de entrantes',
            cached_excel(result),
            f'NIP_ENTRANTES_{year}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type='primary',
        )
        st.caption('O arquivo gerado reúne as obras entrantes e as abas auxiliares de conferência.')

"""NIP | Comparador de Levantamentos - dashboard Streamlit."""
from __future__ import annotations
import base64
import html
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from comparador import analyze, daily_counts, monthly_forecast, make_excel, dt, COL_A, COL_F, COL_O, COL_AA

st.set_page_config(page_title='NIP | Entrantes de Levantamentos',page_icon='📊',layout='wide',initial_sidebar_state='collapsed')
ROOT=Path(__file__).resolve().parent
logo=base64.b64encode((ROOT/'assets/logo_nip_original.jpg').read_bytes()).decode('ascii')

st.markdown('''<style>
:root{--navy:#143b54;--teal:#0b7785;--ink:#193a50;--line:#d5e3ed}
html,body,[data-testid="stAppViewContainer"]{background:#edf2f6!important}
[data-testid="stHeader"]{height:0!important;min-height:0!important;background:transparent!important}
[data-testid="stSidebar"],[data-testid="collapsedControl"],[data-testid="stSidebarCollapsedControl"]{display:none!important}
[data-testid="stMainBlockContainer"],.block-container{width:100%!important;max-width:none!important;padding:0 2.1rem 2rem!important}
[data-testid="stMainBlockContainer"]{padding-top:0!important}
[data-testid="stVerticalBlock"]{gap:.65rem}
.nip-head{background:#143b54;display:flex;align-items:center;gap:22px;padding:12px 25px;margin:0 -2.1rem 12px;min-height:91px;color:#fff}
.nip-head img{width:115px;height:69px;object-fit:contain;mix-blend-mode:screen}
.nip-head-title{font-size:19px;font-weight:800;color:white;letter-spacing:-.2px}
.nip-head-sub{font-size:11px;color:#d4e6ee;margin-top:5px}
.nip-head-right{margin-left:auto;background:#0d8590;padding:10px 15px;border-radius:5px;font-size:12px;font-weight:700;white-space:nowrap}
.nip-rail{font-size:14px;font-weight:800;color:#133b54;margin:6px 0 13px}
.nip-sub{font-size:11px;color:#637a88;margin:7px 0 13px}
.nip-rail-foot{border-top:1px solid #dce5eb;padding-top:20px;margin-top:25px;color:#718995;font-size:11px}
.nip-card{background:white;border:1px solid #d5e3ed;border-radius:9px;padding:15px 17px;margin-bottom:10px;min-height:105px}
.nip-card small{display:block;color:#638092;font-size:11px}
.nip-card strong{display:block;color:#153d59;font-size:26px;margin-top:8px;font-weight:800}
.nip-card em{display:block;font-size:11px;color:#6b8190;font-style:normal;margin-top:4px}
.nip-card.teal{border-bottom:3px solid #0d7785}
.nip-card.blue{border-bottom:3px solid #3b82f6}
.nip-card.green{border-bottom:3px solid #48bd87}
.nip-panel{background:#fff;border:1px solid #d5e3ed;border-radius:8px;padding:13px 15px;margin:9px 0 13px}
.nip-panel-title{color:#143b54;font-weight:800;font-size:14px;margin:4px 0 12px}
.nip-panel-sub{color:#6a8391;font-size:11px}
.nip-intro{color:#526d7c;font-size:12px;margin:2px 0 7px}
[data-testid="stFileUploader"]{border:1px solid #d5e3ed!important;border-radius:7px;background:#fff;padding:6px}
[data-testid="stFileUploaderDropzone"]{padding:9px!important;min-height:66px!important}
[data-testid="stFileUploaderDropzoneInstructions"]{font-size:11px!important}
[data-testid="stFileUploader"] section button{font-size:11px!important}
[data-testid="stTabs"]{background:transparent}
[data-testid="stTabs"] [role="tablist"]{background:#f6f9fc;border:1px solid #d3e2ec;border-radius:7px;padding:7px;gap:7px;flex-wrap:wrap;height:auto!important;margin-bottom:12px}
[data-testid="stTabs"] button[role="tab"]{background:#fff;border:1px solid #b9d0df;border-radius:6px;color:#173a56;padding:8px 12px;height:37px;font-weight:750;font-size:12px}
[data-testid="stTabs"] button[role="tab"][aria-selected="true"]{background:#0d7785!important;border-color:#0d7785!important;color:white!important}
[data-testid="stTabs"] [data-baseweb="tab-highlight"],[data-testid="stTabs"] [data-baseweb="tab-border"]{display:none!important}
[data-testid="stDataFrame"],div[data-testid="stPlotlyChart"]{border:1px solid #d7e4ed;border-radius:8px;background:white;padding:5px}
[data-testid="stMetric"]{border:1px solid #d9e6ed;padding:13px;background:#fff;border-radius:8px}
[data-testid="stAlert"]{border-radius:7px}
[data-testid="stVerticalBlockBorderWrapper"]{border-radius:8px}
@media(max-width:900px){[data-testid="stMainBlockContainer"],.block-container{padding:0 12px 20px!important}.nip-head{margin:0 -12px 12px;padding:10px 15px;gap:12px}.nip-head img{width:70px;height:45px}.nip-head-right{display:none}.nip-head-title{font-size:16px}}
</style>''',unsafe_allow_html=True)

st.markdown(f'''<div class="nip-head"><img src="data:image/jpeg;base64,{logo}" alt="NIP Grupo Ígneo"><div><div class="nip-head-title">NIP | Comparador de Levantamentos</div><div class="nip-head-sub">Obras entrantes · Comparação de bases · Análise diária · Projeção mensal</div></div><div class="nip-head-right">✓ Comparação local de arquivos Excel</div></div>''',unsafe_allow_html=True)

@st.cache_data(show_spinner='Comparando as planilhas...',max_entries=2)
def cached_analysis(old:bytes,new:bytes,year:int):return analyze(old,new,year=year)

@st.cache_data(show_spinner='Preparando arquivo...',max_entries=2)
def cached_excel(result:dict):return make_excel(result)

left,right=st.columns([1.07,4.93],gap='medium',vertical_alignment='top')
with left:
    st.markdown('<div class="nip-rail">Filtros operacionais</div>',unsafe_allow_html=True)
    previous=st.file_uploader('1 · Levantamentos anterior',type=['xlsx'],key='previous',help='Arquivo anterior (base de comparação).')
    current=st.file_uploader('2 · Levantamentos atual',type=['xlsx'],key='current',help='Arquivo mais recente.')
    year=st.number_input('Ano do recebimento (AA)',min_value=2020,max_value=2100,value=2026,step=1)
    st.markdown('<div class="nip-sub">Colunas: <b>A</b> · ID; <b>F</b> · Nota CCS; <b>O</b> · Tipo de Projeto; <b>AA</b> · Data do recebimento.</div>',unsafe_allow_html=True)
    st.markdown('<div class="nip-rail-foot">NIP · Grupo Ígneo<br>Arquivos processados na sessão Streamlit.</div>',unsafe_allow_html=True)

with right:
    tabs=st.tabs(['Dashboard','Entrantes por dia','Obras entrantes','Auditoria','Exportação'])
    if not previous or not current:
        with tabs[0]:
            st.markdown('<div class="nip-panel-title">◉ Resumo da comparação</div>',unsafe_allow_html=True)
            st.info('Envie as duas versões da planilha Levantamentos nos filtros à esquerda para iniciar a comparação.')
            x,y,z=st.columns(3)
            for col,title,description in [(x,'01 · Comparação','Identificar obras que entraram na base atual.'),(y,'02 · Tendência','Estimar o volume mensal pelos entrantes diários.'),(z,'03 · Exportação','Gerar Excel com somente as novas obras.')]:
                with col:st.markdown(f'<div class="nip-card teal"><small>{title}</small><strong style="font-size:15px">{description}</strong></div>',unsafe_allow_html=True)
        with tabs[1]:st.info('Envie os arquivos para ver os entrantes por dia.')
        with tabs[2]:st.info('Envie os arquivos para pesquisar as novas obras.')
        with tabs[3]:st.info('Envie os arquivos para verificar duplicidades e alterações.')
        with tabs[4]:st.info('Envie os arquivos para gerar a terceira planilha.')
        st.stop()
    if previous.getvalue()==current.getvalue():st.warning('Os arquivos são idênticos. Nenhuma diferença será identificada.')
    try:result=cached_analysis(previous.getvalue(),current.getvalue(),int(year))
    except Exception as exc:
        st.error(f'Erro ao analisar os arquivos: {exc}')
        st.stop()
    entrants=result['entrants']; counts=daily_counts(entrants); forecast=monthly_forecast(entrants)
    def frame_rows(rows,with_reason=False):
        headers=result['headers']; return pd.DataFrame(([num,reason,*row] if with_reason else [num,*row] for num,row,reason in rows),columns=['Linha original']+(['Detalhe'] if with_reason else [])+headers)
    def card(label,value,sub='',variant='teal'):
        return f'<div class="nip-card {variant}"><small>{html.escape(label)}</small><strong>{value:,}</strong><em>{html.escape(sub)}</em></div>'.replace(',', '.')

    with tabs[0]:
        st.markdown('<div class="nip-panel-title">◉ Resumo de entrantes</div>',unsafe_allow_html=True)
        cards=st.columns(5,gap='small')
        fields=[('Base anterior',result['old_count'],'Registros analisados','blue'),('Base atual',result['new_count'],'Registros analisados','blue'),('Novos entrantes',len(entrants),f'AA em {year}','green'),('Obras alteradas',len(result['changes']),'Não contam como novas','teal'),('Projeção mensal',forecast['projection'],'Estimativa para 30 dias','teal')]
        for col,(label,value,sub,variant) in zip(cards,fields):
            with col:st.markdown(card(label,value,sub,variant),unsafe_allow_html=True)
        chartcol,statscol=st.columns([1.5,1],gap='small')
        with chartcol:
            st.markdown('<div class="nip-panel-title">▥ Evolução diária de novas obras</div>',unsafe_allow_html=True)
            if counts:
                days=pd.date_range(min(counts),max(counts),freq='D')
                daily=pd.DataFrame({'Data':days,'Entrantes':[counts.get(day.date(),0) for day in days]})
                fig=px.bar(daily,x='Data',y='Entrantes',color_discrete_sequence=['#0d7785'])
                fig.update_layout(plot_bgcolor='white',paper_bgcolor='white',font_color='#143b54',margin=dict(l=8,r=8,t=12,b=5),height=300)
                st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
            else:st.info(f'Nenhum entrante com data AA em {year}.')
        with statscol:
            st.markdown('<div class="nip-panel-title">◉ Indicadores de tendência</div>',unsafe_allow_html=True)
            a,b=st.columns(2)
            a.metric('Média diária',f"{forecast['avg']:.2f}".replace('.',','))
            b.metric('Dias observados',forecast['days'])
            st.metric('Entrantes na janela',forecast['total'])
            if forecast['period']:
                start,end=forecast['period'];st.caption(f'Período das datas AA: {start:%d/%m/%Y} a {end:%d/%m/%Y}')
            st.caption(forecast['note'])
        st.markdown('<div class="nip-panel-title">▥ Distribuição por tipo de projeto (coluna O)</div>',unsafe_allow_html=True)
        if entrants:
            serie=pd.Series([str(row[COL_O]).strip() if row[COL_O] is not None else 'Não informado' for _,row,_ in entrants]).value_counts().head(15)
            fig=px.bar(x=serie.values,y=serie.index,orientation='h',labels={'x':'Entrantes','y':'Tipo de projeto'},color_discrete_sequence=['#0d7785'])
            fig.update_layout(yaxis={'autorange':'reversed'},height=max(220,len(serie)*30+55),plot_bgcolor='white',paper_bgcolor='white',margin=dict(l=10,r=20,t=5,b=5))
            st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
        for issue in result['issues']:st.warning(issue)

    with tabs[1]:
        st.markdown('<div class="nip-panel-title">▥ Entradas diárias e consolidação mensal</div>',unsafe_allow_html=True)
        if counts:
            diary=pd.DataFrame([{'Data AA':day,'Entrantes':qty} for day,qty in sorted(counts.items())]);diary['Mês']=diary['Data AA'].apply(lambda d:d.strftime('%m/%Y'))
            months=sorted(diary['Mês'].unique(),key=lambda v:(int(v[3:]),int(v[:2])),reverse=True)
            chosen=st.selectbox('Mês para detalhamento',months)
            filtered=diary[diary['Mês']==chosen]
            st.metric('Entrantes no mês selecionado',int(filtered['Entrantes'].sum()))
            st.dataframe(filtered[['Data AA','Entrantes']],hide_index=True,use_container_width=True,height=340)
            st.markdown('<div class="nip-panel-title">Consolidado mensal</div>',unsafe_allow_html=True)
            st.dataframe(diary.groupby('Mês',as_index=False)['Entrantes'].sum(),hide_index=True,use_container_width=True)
            st.caption('O consolidado considera somente as obras que surgiram entre os dois arquivos, por data AA.')
        else:st.info('Sem entrantes para o ano escolhido.')

    with tabs[2]:
        st.markdown('<div class="nip-panel-title">▥ Relação de novas obras</div>',unsafe_allow_html=True)
        term=st.text_input('Pesquisar ID, Nota CCS ou tipo de projeto',placeholder='Número da obra, CCS ou tipo...')
        dates=sorted({dt(row[COL_AA]) for _,row,_ in entrants if dt(row[COL_AA])})
        selected=st.multiselect('Filtrar datas AA',dates,format_func=lambda day:day.strftime('%d/%m/%Y'))
        filtered=[]
        for item in entrants:
            _,row,_=item
            if selected and dt(row[COL_AA]) not in selected:continue
            if term and not any(term.casefold() in str(row[i] or '').casefold() for i in (COL_A,COL_F,COL_O)):continue
            filtered.append(item)
        st.caption(f'{len(filtered):,} obras exibidas · {len(entrants):,} entrantes identificados'.replace(',','.'))
        if filtered:
            frame=frame_rows(filtered);names=result['headers']
            st.dataframe(frame[['Linha original',names[COL_A],names[COL_F],names[COL_O],names[COL_AA]]].head(10000),hide_index=True,use_container_width=True,height=470)
            with st.expander('Prévia com todas as colunas originais'):
                st.dataframe(frame.head(200),hide_index=True,use_container_width=True)
            st.download_button('Baixar entrantes filtrados (.xlsx)',make_excel(result,filtered),'NIP_ENTRANTES_FILTRADOS.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:st.info('Nenhuma obra corresponde aos filtros.')

    with tabs[3]:
        st.markdown('<div class="nip-panel-title">◉ Auditoria da comparação</div>',unsafe_allow_html=True)
        columns=st.columns(3)
        for col,label,key in zip(columns,['Duplicados atuais','Novos fora do ano','Sem identificação / data'],['duplicates','excluded','invalid']):col.metric(label,len(result[key]))
        st.caption('Obras existentes cujos campos A, F, O ou AA mudaram são registradas como alterações, não entrantes.')
        with st.expander('Registros existentes alterados'):
            st.dataframe(frame_rows(result['changes']).head(500) if result['changes'] else pd.DataFrame(),hide_index=True,use_container_width=True)
        with st.expander('Registros fora do ano, duplicidades e inconsistências'):
            records=result['excluded']+result['invalid']+result['duplicates']
            st.dataframe(frame_rows(records,with_reason=True).head(500) if records else pd.DataFrame(),hide_index=True,use_container_width=True)
        for issue in result['issues']:st.warning(issue)

    with tabs[4]:
        st.markdown('<div class="nip-panel-title">▤ Exportação da terceira planilha</div>',unsafe_allow_html=True)
        st.markdown(f'<div class="nip-card green"><small>Obras novas com recebimento em {year}</small><strong>{len(entrants):,}</strong><em>Arquivo Excel com colunas originais e resumo de auditoria</em></div>'.replace(',','.'),unsafe_allow_html=True)
        st.download_button('⬇ Baixar planilha completa de entrantes',cached_excel(result),f'NIP_ENTRANTES_{year}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
        st.caption('A aba OBRAS ENTRANTES contém as obras novas; as demais abas documentam totais, entradas diárias e auditoria.')

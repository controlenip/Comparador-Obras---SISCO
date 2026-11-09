from __future__ import annotations
from datetime import date
import html
import pandas as pd
import plotly.express as px
import streamlit as st
from comparador import analyze, daily_counts, monthly_forecast, make_excel, dt, COL_A, COL_F, COL_O, COL_AA

st.set_page_config(page_title='NIP | Entrantes de Levantamentos',page_icon='📊',layout='wide',initial_sidebar_state='expanded')

st.markdown('''<style>
:root{--navy:#143b54;--teal:#0d7785}
[data-testid="stAppViewContainer"]{background:#edf2f6}
[data-testid="stSidebar"]{background:#fff;border-right:1px solid #dbe4ea}
[data-testid="stHeader"]{background:#143b54}
.block-container{padding-top:1.3rem;max-width:1600px}
.nip-header{background:#143b54;color:#fff;padding:14px 19px;border-radius:8px;display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
.nip-header h1{font-size:23px;color:white;margin:0}
.nip-header small{color:#d0e1ec;font-size:12px}
.kpanel{background:#fff;border-radius:7px;border-bottom:4px solid #0d7785;padding:16px;margin-bottom:8px;min-height:112px}
.kpanel span{display:block;color:#607c8e;font-size:12px;font-weight:600}.kpanel strong{display:block;color:#143b54;font-size:28px;margin-top:8px}
.section-title{font-size:18px;font-weight:800;color:#143b54;margin:12px 0}
[data-testid="stMetric"]{background:white;padding:14px;border-radius:8px;border-bottom:3px solid #0d7785}
div[data-testid="stFileUploader"]{background:#fff;border:1px solid #dae6ec;border-radius:8px;padding:5px}
</style>''',unsafe_allow_html=True)

@st.cache_data(show_spinner='Comparando as planilhas... ',max_entries=3)
def cached_analysis(old: bytes,new: bytes,year:int): return analyze(old,new,year=year)

@st.cache_data(show_spinner='Gerando Excel...',max_entries=3)
def cached_excel(result:dict): return make_excel(result)

with st.sidebar:
    st.image('assets/logo_nip.png',width=185)
    st.markdown('### Configuração da análise')
    previous=st.file_uploader('1 · Planilha anterior',type=['xlsx'],key='previous',help='Arquivo Levantamentos do dia anterior ou snapshot mais antigo.')
    current=st.file_uploader('2 · Planilha atual',type=['xlsx'],key='current',help='Arquivo Levantamentos mais recente.')
    year=st.number_input('Ano da data AA',min_value=2020,max_value=2100,value=2026,step=1)
    st.caption('Comparação: A (ID), F (Nota CCS), O (tipo de projeto), AA (data do recebimento).')
    st.divider()
    st.caption('NIP · Grupo Ígneo | Dados processados na sessão Streamlit')

st.markdown('<div class="nip-header"><div><h1>NIP | Entrantes de Levantamentos</h1><small>Comparação de bases · Entradas diárias · Projeção mensal</small></div><span style="font-size:30px">▤</span></div>',unsafe_allow_html=True)

if not previous or not current:
    st.info('Envie duas planilhas na barra lateral para identificar os entrantes. Selecione primeiro a versão anterior e depois a versão atual.')
    a,b,c=st.columns(3)
    for col,n,title,detail in [(a,'01','Comparar','Localizar registros presentes apenas na base atual.'),(b,'02','Analisar','Contabilizar entradas por dia e por tipo de projeto.'),(c,'03','Exportar','Baixar planilha contendo somente as obras novas.')]:
        with col:st.markdown(f'<div class="kpanel"><span>{n} · {title}</span><strong style="font-size:17px">{detail}</strong></div>',unsafe_allow_html=True)
    st.stop()
if previous.getvalue()==current.getvalue():
    st.warning('Os dois arquivos são idênticos. Envie versões diferentes para localizar alterações.')
try:
    result=cached_analysis(previous.getvalue(),current.getvalue(),int(year))
except Exception as exc:
    st.error(f'Não foi possível analisar as planilhas: {exc}')
    st.stop()

for warning in result['issues']:st.warning(warning)
entrants=result['entrants']
counts=daily_counts(entrants)
forecast=monthly_forecast(entrants)

k=st.columns(5)
for column,name,val,subtitle in zip(k,['Base anterior','Base atual','Obras entrantes','Alterações existentes','Projeção / 30 dias'],
    [result['old_count'],result['new_count'],len(entrants),len(result['changes']),forecast['projection']],
    ['linhas lidas','linhas lidas',f'AA em {year}','A/F/O/AA','estimativa indicativa']):
    with column:
        st.markdown(f'<div class="kpanel"><span>{name}</span><strong>{val:,}</strong><span>{subtitle}</span></div>'.replace(',', '.'),unsafe_allow_html=True)

tabs=st.tabs(['📊 Visão geral','📅 Entrantes por dia','📋 Obras entrantes','🔎 Auditoria e exportação'])

def df_rows(rows, with_reason=False):
    cols=result['headers']
    frame=pd.DataFrame(([num,reason,*r] if with_reason else [num,*r] for num,r,reason in rows),
        columns=['Linha original']+(['Detalhe'] if with_reason else [])+cols)
    return frame

with tabs[0]:
    st.markdown('<div class="section-title">Análise do volume de novas demandas</div>',unsafe_allow_html=True)
    x,y=st.columns([1.65,1])
    with x:
        st.markdown('#### Evolução diária dos entrantes')
        if counts:
            all_days=pd.date_range(min(counts),max(counts),freq='D')
            dfd=pd.DataFrame({'Data':all_days,'Entrantes':[counts.get(d.date(),0) for d in all_days]})
            fig=px.bar(dfd,x='Data',y='Entrantes',color_discrete_sequence=['#0d7785'])
            fig.update_layout(plot_bgcolor='white',paper_bgcolor='white',margin=dict(l=8,r=8,t=14,b=5),height=320)
            st.plotly_chart(fig,use_container_width=True)
        else:st.info('Nenhuma obra nova no ano filtrado.')
    with y:
        st.markdown('#### Indicadores da projeção')
        st.metric('Média observada de entradas/dia',f"{forecast['avg']:.2f}".replace('.',','))
        st.metric('Entradas na janela',forecast['total'])
        st.metric('Dias considerados',forecast['days'])
        if forecast['period']:
            start,end=forecast['period'];st.caption(f'Janela AA: {start:%d/%m/%Y} a {end:%d/%m/%Y}')
        st.caption(forecast['note'])
    if entrants:
        st.markdown('#### Distribuição por tipo de projeto (coluna O)')
        by_type=pd.Series([(str(row[COL_O]).strip() if row[COL_O] is not None else 'Não informado') for _,row,_ in entrants]).value_counts().head(15)
        st.bar_chart(by_type,color='#0d7785',horizontal=True)

with tabs[1]:
    st.markdown('<div class="section-title">Entradas por data da coluna AA</div>',unsafe_allow_html=True)
    if counts:
        diary=pd.DataFrame([{'Data AA':day,'Entrantes':qty} for day,qty in sorted(counts.items())])
        diary['Mês']=diary['Data AA'].apply(lambda d:d.strftime('%m/%Y'))
        months=sorted(diary['Mês'].unique(),key=lambda v:(int(v[3:]),int(v[:2])),reverse=True)
        chosen=st.selectbox('Mês para detalhamento',months)
        filtered=diary[diary['Mês']==chosen]
        total=int(filtered['Entrantes'].sum())
        st.metric('Total de entrantes no mês selecionado',total)
        st.dataframe(filtered[['Data AA','Entrantes']],hide_index=True,use_container_width=True)
        month_summary=diary.groupby('Mês',as_index=False)['Entrantes'].sum()
        st.markdown('#### Consolidado mensal, por data AA')
        st.dataframe(month_summary.sort_values('Mês',ascending=False),hide_index=True,use_container_width=True)
        st.caption('Apenas obras que surgiram entre os dois arquivos. Não equivale ao total histórico de obras recebidas no mês.')
    else:st.info('Sem registros para exibir.')

with tabs[2]:
    st.markdown('<div class="section-title">Relação de novas obras</div>',unsafe_allow_html=True)
    term=st.text_input('Pesquisar por ID, Nota CCS ou Tipo de Projeto',placeholder='Ex.: 111...',key='term')
    dates=sorted({dt(r[COL_AA]) for _,r,_ in entrants if dt(r[COL_AA])})
    selected_dates=st.multiselect('Filtrar datas de recebimento AA',dates,format_func=lambda d:d.strftime('%d/%m/%Y'))
    filtered=[]
    for item in entrants:
        _,r,_=item
        if selected_dates and dt(r[COL_AA]) not in selected_dates:continue
        if term and not any(term.casefold() in str(r[i] or '').casefold() for i in (COL_A,COL_F,COL_O)):continue
        filtered.append(item)
    st.caption(f'{len(filtered):,} obras exibidas · {len(entrants):,} entrantes identificados'.replace(',','.'))
    if filtered:
        frame=df_rows(filtered)
        selected=['Linha original',result['headers'][COL_A],result['headers'][COL_F],result['headers'][COL_O],result['headers'][COL_AA]]
        st.dataframe(frame[selected].head(10000),hide_index=True,use_container_width=True,height=440)
        if len(frame)>10000:st.caption('Exibição limitada às primeiras 10 mil linhas; a exportação contém todas as linhas filtradas.')
        with st.expander('Ver todas as colunas originais (prévia de 200 linhas)'):
            st.dataframe(frame.head(200),hide_index=True,use_container_width=True)
        st.download_button('⬇️ Baixar obras filtradas (.xlsx)',make_excel(result,filtered),'NIP_OBRAS_ENTRANTES_FILTRADAS.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:st.info('Nenhum registro corresponde aos filtros.')

with tabs[3]:
    st.markdown('<div class="section-title">Auditoria da comparação</div>',unsafe_allow_html=True)
    c1,c2,c3=st.columns(3)
    c1.metric('Duplicados no arquivo atual',len(result['duplicates']))
    c2.metric('Novas fora do ano escolhido',len(result['excluded']))
    c3.metric('Identificadores/datas inválidos',len(result['invalid']))
    st.caption('Registros que já existiam no arquivo anterior mas tiveram alteração em A, F, O ou AA não são considerados entradas novas.')
    with st.expander('Registros existentes alterados'):
        st.dataframe(df_rows(result['changes']).head(500) if result['changes'] else pd.DataFrame(),hide_index=True,use_container_width=True)
    with st.expander('Registros fora do ano e inconsistências'):
        st.dataframe(df_rows(result['excluded']+result['invalid']+result['duplicates'],with_reason=True).head(500) if (result['excluded'] or result['invalid'] or result['duplicates']) else pd.DataFrame(),hide_index=True,use_container_width=True)
    st.markdown('#### Exportação completa')
    st.download_button('⬇️ Baixar terceira planilha — apenas entrantes + auditoria',cached_excel(result),f'NIP_ENTRANTES_{year}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
    st.caption('A aba OBRAS ENTRANTES contém apenas novas obras, preservando as colunas originais. As demais abas documentam totais e verificações.')

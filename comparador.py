"""NIP - comparacao de snapshots Levantamentos.xlsx, sem enviar dados a servidores externos."""
from __future__ import annotations
from collections import Counter
from datetime import date, datetime, timedelta
from io import BytesIO
from typing import Any
import calendar
import re
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

COL_A, COL_F, COL_O, COL_AA = 0, 5, 14, 26


def norm(value: Any) -> str:
    if value is None: return ""
    if isinstance(value, (datetime, date)): return value.isoformat()
    if isinstance(value, float) and value.is_integer(): return str(int(value))
    return re.sub(r"\s+", " ", str(value).strip()).upper()


def dt(value: Any) -> date | None:
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    if isinstance(value, (float, int)):
        if 30000 < value < 90000:
            try: return openpyxl.utils.datetime.from_excel(value).date()
            except (ValueError, OverflowError): return None
        return None
    if value is None: return None
    s = str(value).strip()
    if not s: return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try: return datetime.strptime(s[:19] if ' %H:' in fmt else s[:10],fmt).date()
        except ValueError: pass
    try: return datetime.fromisoformat(s).date()
    except ValueError: return None


def read_sheet(source: bytes, sheet_name: str | None = None):
    wb = openpyxl.load_workbook(BytesIO(source), read_only=True, data_only=True)
    try:
        if sheet_name and sheet_name not in wb.sheetnames: raise ValueError(f"Aba '{sheet_name}' nao encontrada.")
        ws = wb[sheet_name] if sheet_name else (wb['Levantamentos'] if 'Levantamentos' in wb.sheetnames else wb.active)
        ws.reset_dimensions()
        it = ws.iter_rows(values_only=True)
        header = next(it, None)
        if header is None or len(header) <= COL_AA: raise ValueError('A aba precisa ter pelo menos 27 colunas (ate AA).')
        names = [str(x).strip() if x is not None else f'Coluna {get_column_letter(i+1)}' for i,x in enumerate(header)]
        records = []
        for excel_row, values in enumerate(it, 2):
            vals = tuple(values) + (None,)*max(0,len(names)-len(values))
            if all(v is None or str(v).strip()=='' for v in vals): continue
            records.append((excel_row,vals[:len(names)]))
        return names,records,ws.title
    finally: wb.close()


def key(row: tuple):
    """A e a chave principal. Sem A: F+O; sem ambos: ignora da comparacao."""
    a,f,o=norm(row[COL_A]),norm(row[COL_F]),norm(row[COL_O])
    if a: return ('A',a)
    if f: return ('F+O',f,o)
    return None


def analyze(old: bytes, new: bytes, year: int = 2026, old_sheet: str | None = None, new_sheet: str | None = None):
    h0,r0,s0=read_sheet(old,old_sheet)
    h1,r1,s1=read_sheet(new,new_sheet)
    issues=[]
    for idx in (COL_A,COL_F,COL_O,COL_AA):
        if norm(h0[idx])!=norm(h1[idx]):
            issues.append(f'Cabecalho diferente na coluna {get_column_letter(idx+1)}: "{h0[idx]}" x "{h1[idx]}". Conferir se os arquivos usam o mesmo modelo.')
    prev={}; dup_old=0
    previous_fo={}
    previous_fo_counts=Counter()
    current_fo_counts=Counter((norm(row[COL_F]),norm(row[COL_O])) for _,row in r1 if norm(row[COL_F]))
    for num,row in r0:
        k=key(row)
        if k is None: continue
        if k in prev: dup_old+=1
        else: prev[k]=(num,row)
        fo=(norm(row[COL_F]),norm(row[COL_O]))
        if fo[0]:
            previous_fo_counts[fo]+=1
            previous_fo[fo]=(num,row)
    entrants=[]; changes=[]; excluded=[]; invalid=[]; dup_new=[]; seen={}
    for num,row in r1:
        k=key(row)
        d=dt(row[COL_AA])
        if k is None:
            invalid.append((num,row,'Sem ID (A) e sem Nota CCS (F)'))
            continue
        if k in seen:
            dup_new.append((num,row,'Identificador repetido na planilha atual'))
            continue
        seen[k]=num
        prior=prev.get(k)
        if prior is None:
            fo=(norm(row[COL_F]),norm(row[COL_O]))
            if fo[0] and previous_fo_counts[fo]==1 and current_fo_counts[fo]==1:
                prior=previous_fo[fo]
        if prior is not None:
            previous=prior[1]
            changed=[get_column_letter(i+1) for i in (COL_A,COL_F,COL_O,COL_AA) if norm(previous[i]) != norm(row[i])]
            if changed: changes.append((num,row,', '.join(changed)))
            continue
        if not d:
            invalid.append((num,row,'AA vazia ou data invalida'))
        elif d.year!=year:
            excluded.append((num,row,f'Ano de AA diferente de {year}'))
        else: entrants.append((num,row,''))
    if any('A' in reason.split(', ') for _,_,reason in changes):issues.append('Ha IDs de levantamento alterados com mesma Nota CCS e Tipo de Projeto (F+O) unicos; tratados como alteracao, nao como entrante.')
    if dup_old:issues.append(f'{dup_old} identificadores repetidos na base anterior; foi usada a primeira ocorrencia de cada identificador.')
    if dup_new:issues.append(f'{len(dup_new)} identificadores repetidos na base atual; foi usada a primeira ocorrencia de cada identificador.')
    if invalid:issues.append(f'{len(invalid)} registros na base atual sem chave ou com AA invalida, nao contabilizados como entrantes.')
    if excluded:issues.append(f'{len(excluded)} registros novos com AA fora de {year}, nao contabilizados como entrantes.')
    return {'headers':h1,'old_count':len(r0),'new_count':len(r1),'entrants':entrants,'changes':changes,
            'excluded':excluded,'invalid':invalid,'duplicates':dup_new,'issues':issues,'old_sheet':s0,'new_sheet':s1}


def daily_counts(items):
    return Counter(dt(r[COL_AA]) for _,r,_ in items if dt(r[COL_AA]))


def monthly_forecast(items, ref_date: date | None = None, window: int = 30):
    """Media de entradas/dia com dias sem registro no periodo inclusos; projecao 30 dias.
    ref_date: ultimo dia de cadastro AA visto nos entrantes, nao a data do upload.
    """
    counts=daily_counts(items)
    if not counts: return {'avg':0.,'projection':0,'days':0,'period':None,'total':0,'note':'Sem entrantes datados para estimar.'}
    latest=max(counts)
    end=ref_date or latest
    start=end-timedelta(days=max(1,window)-1)
    total=sum(n for day,n in counts.items() if start<=day<=end)
    # Denominador inclui todos os dias, mas nao anteriores a primeira data disponivel.
    first=max(start,min(counts))
    days=max(1,(end-first).days+1)
    avg=total/days
    return {'avg':avg,'projection':round(avg*30),'days':days,'period':(first,end),'total':total,
            'note':'Projecao indicativa de 30 dias baseada nas datas AA dos entrantes; nao e a demanda total do mes. A contagem e limitada aos dois snapshots enviados.'}


def _safe_excel(value):
    if isinstance(value,(list,dict)): value=str(value)
    if isinstance(value,str) and value.startswith(('=', '+', '-', '@')): return "'"+value
    return value


def make_excel(result:dict, filter_items=None):
    """Excel com 100% das colunas originais dos entrantes e abas de apoio."""
    records=result['entrants'] if filter_items is None else filter_items
    wb=openpyxl.Workbook(); ws=wb.active; ws.title='OBRAS ENTRANTES'
    headers=result['headers']
    ws.append(['Linha original'] + headers)
    for num,row,_ in records: ws.append([num]+[_safe_excel(v) for v in row])
    st=wb.create_sheet('RESUMO');st.append(['Indicador','Quantidade'])
    for label,k in [('Linhas base anterior','old_count'),('Linhas base atual','new_count')]:st.append([label,result[k]])
    for label,k in [('Entrantes 2026','entrants'),('Registros alterados','changes'),('Novos fora do ano','excluded'),('Novos com problemas','invalid'),('Duplicados atuais','duplicates')]:st.append([label,len(result[k])])
    st.append(['Observacao','Entrante significa ausente no snapshot anterior, presente no atual e AA no ano selecionado.'])
    d=wb.create_sheet('ENTRANTES POR DIA');d.append(['Data AA','Quantidade'])
    for day,n in sorted(daily_counts(records).items()):d.append([day,n])
    for row in d.iter_rows(min_row=2,min_col=1,max_col=1):row[0].number_format='DD/MM/YYYY'
    for tab,keyname in [('ALTERACOES','changes'),('FORA DO ANO','excluded'),('INCONSISTENCIAS','invalid'),('DUPLICADOS','duplicates')]:
        t=wb.create_sheet(tab);t.append(['Linha original','Motivo/colunas']+headers)
        for num,row,reason in result[keyname]:t.append([num,reason]+[_safe_excel(v) for v in row])
    for sheet in wb:
        sheet.freeze_panes='B2' if sheet==ws else 'A2'
        sheet.sheet_view.showGridLines=False
        sheet.auto_filter.ref=sheet.dimensions
        for cell in sheet[1]:cell.fill=PatternFill('solid',fgColor='143B54');cell.font=Font(color='FFFFFF',bold=True);cell.alignment=Alignment(vertical='center')
        sheet.row_dimensions[1].height=28
        for j in range(1,sheet.max_column+1):
            hdr=str(sheet.cell(1,j).value or '')
            sheet.column_dimensions[get_column_letter(j)].width=min(37,max(15,len(hdr)+4))
        for row in sheet.iter_rows(min_row=2):
            if row[0].row%2==0:
                for c in row:c.fill=PatternFill('solid',fgColor='F0F5F8')
    out=BytesIO();wb.save(out);return out.getvalue()

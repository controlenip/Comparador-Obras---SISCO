import sys
from pathlib import Path
from datetime import datetime
from io import BytesIO
import openpyxl
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from comparador import analyze,monthly_forecast,make_excel

def workbook(rows):
    w=openpyxl.Workbook();s=w.active;s.title='Levantamentos'
    hdr=[f'Col {i}' for i in range(1,28)];hdr[0]='ID do levantamento';hdr[5]='Nota CCS';hdr[14]='Tipo de Projeto(PI)';hdr[26]='Recebimento(data criação)';s.append(hdr)
    for a,f,o,dt in rows:
        row=[None]*27;row[0]=a;row[5]=f;row[14]=o;row[26]=dt;s.append(row)
    b=BytesIO();w.save(b);return b.getvalue()

def test_new_old_filter_and_change():
    old=workbook([('1','111','AA',datetime(2025,1,1)),('2','222','BB',datetime(2026,1,1))])
    new=workbook([('1','111','AA',datetime(2025,1,1)),('2','222','CC',datetime(2026,1,1)),('3','333','AA',datetime(2026,11,2)),('4','444','AA',datetime(2025,11,2)),('3','333','AA',datetime(2026,11,2))])
    r=analyze(old,new)
    assert len(r['entrants'])==1
    assert len(r['changes'])==1
    assert len(r['excluded'])==1
    assert len(r['duplicates'])==1
    assert r['entrants'][0][1][0]=='3'
    b=make_excel(r)
    w=openpyxl.load_workbook(BytesIO(b),read_only=True)
    assert w['OBRAS ENTRANTES'].max_row==2

def test_no_data():
    b=workbook([('1','1','A',datetime(2026,1,1))]);r=analyze(b,b)
    assert len(r['entrants'])==0
    assert monthly_forecast(r['entrants'])['projection']==0

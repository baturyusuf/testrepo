"""Regression, formatting and physical page-fit checks for the browser workbook."""
import math
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import sys
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import adaptive_layout as layout
import web_bridge


def test_all_combinations():
    checks=0
    for weeks in range(10,21):
        for marks in range(1,6):
            for vize,final in [(8,min(16,weeks)), (1,weeks), (weeks-1,weeks), (2,min(9,weeks))]:
                for min_mark, name_max in [(8,58),(15,58),(20,30),(20,80)]:
                    fixed=(layout.SERIAL_WIDTH_MM,layout.STUDENT_NUMBER_WIDTH_MM,name_max)
                    pages=layout.page_plan(weeks,marks,vize,final,min_mark,fixed)
                    expanded=[g for page in pages for g in page]
                    assert [g.week for g in expanded]==list(range(1,weeks+1))
                    assert sum(g.kind=='exam' for g in expanded)==2
                    for page in pages:
                        width=layout.mark_width_mm(page,min_mark,fixed)
                        assert min_mark-0.001 <= width <= max(18,min_mark)+0.001
                        assert sum(g.span for g in page)*width+sum(fixed) <= layout.PRINTABLE_MM+0.01
                    checks+=1
    assert len(layout.page_plan(20,5,8,16,15))>=5
    assert len(layout.page_plan(20,5,8,16,8)) < len(layout.page_plan(20,5,8,16,15))
    return checks


def test_names():
    short=[('42','AYŞE YILMAZ'),('43','ALİ CAN')]
    middle=[('42','MUHAMMED YUSUF BATUR')]
    long=[('42','YÜKSEK UZUNLUKTA ÖĞRENCİ İSMİ SOYADI VE ÇOK UZUN EKLER')]
    widths=[layout.estimated_name_width_mm(x) for x in (short,middle,long)]
    assert layout.NAME_MIN_MM<=widths[0]<widths[1]<widths[2]==58, widths
    assert layout.estimated_name_width_mm(long,maximum_mm=42)<=42
    assert layout.estimated_name_width_mm(middle,font_size_pt=12)>widths[1]
    return widths


def test_workbook_and_bridge():
    checked=0
    with TemporaryDirectory() as t:
        work=Path(t)
        for n,count in [(1,1),(2,105)]:
            folder=work/'uploads'/str(n);folder.mkdir(parents=True)
            wb=Workbook();ws=wb.active
            ws.append(['Öğrenci No_42','Adı_42','Soyadı_42','Program_42'])
            for k in range(count):
                name='AYŞE' if k%3 else 'ABDÜLMECİT';last='YILMAZ' if k%3 else 'KARABULUTOĞLU'
                ws.append((f'0{k+1:07d}',name,last,'BİLGİSAYAR MÜHENDİSLİĞİ'))
            wb.save(folder/f'205-b{n}.xlsx')
        web_bridge.WORK=work
        configs=[
            (16,1,8,16,15,8,'Aptos',58),
            (16,3,8,16,10,10,'Arial',58),
            (10,2,4,9,20,14,'Calibri',40),
            (20,5,8,16,15,8,'Times New Roman',58),
            (20,1,6,19,8,6,'Aptos',30),
        ]
        for weeks,signatures,vize,final,minmark,fontpt,family,namemax in configs:
            data={
                'weeks':weeks,'signatures':signatures,'midterm':vize,'final':final,
                'minMarkWidthMm':minmark,'fontSizePt':fontpt,
                'fontFamily':family,'nameMaxWidthMm':namemax,
                'course':'BLG103 - Programlama I','instructor':'Dr. Öğr. Üyesi Deneme',
                'academicYear':'2026-2027','semester':'Güz',
                'files':[{'name':'205-b1.xlsx'},{'name':'205-b2.xlsx'}]
            }
            results=json.loads(web_bridge.generate(json.dumps(data)))
            assert not results['errors'],results['errors']
            assert len(results['results'])==2
            for original,output in zip((1,105),results['results']):
                assert output['count']==original
                assert output['blankRows']==math.ceil(original/17)*17-original
                wb=load_workbook(output['file'])
                assert len(wb.worksheets)==len(output['sheets'])
                all_groups=[g for sh in output['sheets'] for g in sh['groups']]
                assert [g['week'] for g in all_groups]==list(range(1,weeks+1))
                assert sum(g['kind']=='exam' for g in all_groups)==2
                for ws,sh in zip(wb.worksheets,output['sheets']):
                    assert ws.page_setup.orientation=='landscape'
                    assert ws.page_setup.fitToWidth==1
                    assert ws.print_title_rows=='$1:$6'
                    assert len(sh['pages'])==math.ceil(original/17)
                    assert len(ws.row_breaks.brk)==max(0,len(sh['pages'])-1)
                    assert sh['markWidthMm']>=minmark-0.01
                    assert sh['totalWidthMm']<=289.01
                    assert sh['fontSizePt']==fontpt
                    assert sh['fontFamily']==family
                    assert sh['fixedWidthsMm'][0]==7.5
                    assert sh['fixedWidthsMm'][1]==21
                    assert sh['fixedWidthsMm'][2]<=namemax
                    assert sh['fixedWidthsMm'][2]>=layout.NAME_MIN_MM
                    assert 6 <= ws['C7'].font.sz <= fontpt
                    assert ws['C7'].font.sz == sh['pages'][0][0]['nameFontPt']
                    assert ws['C7'].value in ('ABDÜLMECİT KARABULUTOĞLU','AYŞE YILMAZ')
                    assert ws['B7'].value=='00000001'
                    assert ws.row_dimensions[7].height>=layout.MIN_STUDENT_HEIGHT_PT
                    assert ws.row_dimensions[7].height<=layout.MAX_STUDENT_HEIGHT_PT
                wb.close();checked+=1
        for bad in ({'minMarkWidthMm':50},{'fontSizePt':50},{'nameMaxWidthMm':10},{'fontFamily':'not a font'}):
            try:web_bridge.generate(json.dumps({**data,**bad}))
            except ValueError:pass
            else:raise AssertionError('Invalid setting accepted '+str(bad))
    return checked

if __name__=='__main__':
    print('Cases checked:',test_all_combinations(),'layouts;',test_names(),'name widths;',test_workbook_and_bridge(),'workbooks.')
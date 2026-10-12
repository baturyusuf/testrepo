"""Pyodide bridge: create adaptive A4 attendance workbooks and PDF page manifests.

All data remains in the browser's sandbox filesystem. Does not use a server.
"""
import json
import re
import shutil
from pathlib import Path
import adaptive_layout as layout

WORK = Path('/home/pyodide/work')


def _clean_name(name):
    name = str(name).replace('\\', '/').split('/')[-1]
    name = re.sub(r'[^\w.\- ]', '_', name, flags=re.UNICODE).strip(' .')
    return name[:110] or 'obs.xlsx'


def generate(payload):
    settings = json.loads(payload)
    weeks, signatures, midterm, final = layout.validate_settings(
        settings.get('weeks'), settings.get('signatures'),
        settings.get('midterm'), settings.get('final'))
    settings.update(weeks=weeks, signatures=signatures, midterm=midterm, final=final)
    course = ' '.join(str(settings.get('course') or '').split())
    instructor = ' '.join(str(settings.get('instructor') or '').split())
    if not course or not instructor:
        raise ValueError('Ders kodu/adı ve öğretim elemanı gereklidir')
    academic_year = str(settings.get('academicYear', '')).strip()
    semester = str(settings.get('semester', '')).strip()
    match = re.fullmatch(r'(20\d{2})-(20\d{2})', academic_year)
    if not match or int(match.group(2)) != int(match.group(1)) + 1:
        raise ValueError('Eğitim öğretim yılı YYYY-YYYY biçiminde olmalıdır')
    if semester not in ('Güz', 'Bahar'):
        raise ValueError('Dönem Güz veya Bahar olmalıdır')
    settings['course'],settings['instructor']=course,instructor
    output = WORK / 'output'
    if output.exists():
        shutil.rmtree(output)
    output.mkdir(parents=True)
    result={'results':[], 'errors':[], 'configuration':{'weeks':weeks,'signatures':signatures,
           'midterm':midterm,'final':final,'horizontalSections':len(layout.page_plan(weeks,signatures,midterm,final))}}
    for index, submitted in enumerate(settings.get('files', []), 1):
        try:
            name=_clean_name(submitted['name'])
            source=WORK/'uploads'/str(index)/name
            if source.suffix.lower()!='.xlsx':
                raise ValueError('Yalnızca .xlsx dosyaları desteklenmektedir')
            if not source.is_file():
                raise ValueError('Yüklenen dosya bulunamadı')
            dest, students, department, warnings, sheet_data, blanks=layout.create_file(
                source,output/str(index),settings)
            result['results'].append({
                'input':submitted['name'],'file':str(dest),'filename':dest.name,
                'count':len(students),'blankRows':blanks,'department':department,
                'pagesPerSheet':len(sheet_data[0]['pages']),'warnings':warnings,'sheets':sheet_data
            })
        except Exception as exc:
            result['errors'].append({'file':submitted.get('name','(dosya)'),'message':str(exc)})
    return json.dumps(result,ensure_ascii=False)
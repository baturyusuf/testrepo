"""Browser adapter for the original OBS attendance generator.

Only runs in Pyodide. All student data is processed in the browser's WASM FS.
"""
import json
import re
import shutil
from pathlib import Path
from openpyxl import load_workbook
import yoklama_olustur as core

WORK = Path('/home/pyodide/work')


def _clean_name(name):
    name = name.replace('\\', '/').split('/')[-1]
    name = re.sub(r'[^\w.\- ]', '_', name, flags=re.UNICODE).strip(' .')
    return name[:110] or 'obs.xlsx'


def _sheet_data(ws, imza, capacity):
    first = int(ws.print_title_rows.split(':')[-1].strip('$')) + 1
    # One-page outputs have no explicit row break; infer last student row
    # from the (reconstructed) exam merge blocks instead.
    exam_merges = [r for r in ws.merged_cells.ranges if r.max_row >= first]
    last = max(r.max_row for r in exam_merges)
    blocks = {}
    for r in exam_merges:
        blocks.setdefault((r.min_col, r.max_col), r.min_row)
    groups = []
    for cell in ws[5]:
        if isinstance(cell.value, str) and cell.value.upper().startswith('HAFTA '):
            groups.append({'column': cell.column, 'label': cell.value.strip(), 'kind': 'week', 'span': imza})
    for (c1, c2), top in blocks.items():
        # After creating a sheet, each page gets its own vertically merged exam block.
        label = ws.cell(first, c1).value or ws.cell(top, c1).value or 'SINAV'
        groups.append({'column': c1, 'label': str(label), 'kind': 'exam', 'span': c2-c1+1})
    groups.sort(key=lambda group: group['column'])
    pages = []
    for start in range(first, last + 1, capacity):
        rows = []
        for r in range(start, min(start + capacity, last + 1)):
            rows.append({
                'rank': ws.cell(r, 1).value,
                'number': ws.cell(r, 2).value,
                'name': ws.cell(r, 3).value,
            })
        pages.append(rows)
    return {
        'name': ws.title,
        'title': str(ws['A1'].value or ''),
        'course': str(core.bilgi_hucresi(ws, core.DERS_BASLIGI).value or ''),
        'instructor': str(core.bilgi_hucresi(ws, core.HOCA_BASLIGI).value or ''),
        'department': str(ws['A4'].value or ''),
        'groups': groups,
        'capacity': capacity,
        'pages': pages,
    }


def generate(payload):
    """Accept JSON config, return JSON manifest with output paths and PDF layouts."""
    settings = json.loads(payload)
    imza = int(settings['signatures'])
    if imza not in (1, 2, 3):
        raise ValueError('Geçersiz imza sayısı')
    template = WORK / 'templates' / f'Yoklama-{imza}-Sign-Template-Ready-70-Students.xlsx'
    if not template.is_file():
        raise ValueError('Şablon bulunamadı')
    # openpyxl read-only doesn't preserve some layout metadata, but header cells work.
    original_wb = load_workbook(template, read_only=False)
    first = original_wb.worksheets[0]
    capacity = core.sablon_duzeni(first).sayfa_basina
    course = core.temizle(settings.get('course')) or core.temizle(core.bilgi_hucresi(first, core.DERS_BASLIGI).value)
    instructor = core.temizle(settings.get('instructor')) or core.temizle(core.bilgi_hucresi(first, core.HOCA_BASLIGI).value)
    original_wb.close()
    if not course or not instructor:
        raise ValueError('Ders kodu/adı ve öğretim elemanı gereklidir')
    academic_year = str(settings.get('academicYear', '')).strip()
    semester = str(settings.get('semester', '')).strip()
    match = re.fullmatch(r'(20\d{2})-(20\d{2})', academic_year)
    if not match or int(match.group(2)) != int(match.group(1)) + 1:
        raise ValueError('Eğitim öğretim yılı YYYY-YYYY biçiminde olmalıdır')
    if semester not in ('Güz', 'Bahar'):
        raise ValueError('Dönem Güz veya Bahar olmalıdır')
    title = (f'İSTANBUL RUMELİ ÜNİVERSİTESİ {academic_year} EĞİTİM ÖĞRETİM YILI '
             f'{core.tr_buyuk(semester)} DÖNEMİ DERS YOKLAMA LİSTESİ')
    output = WORK / 'output'
    if output.exists():
        shutil.rmtree(output)
    output.mkdir(parents=True)
    manifest = {'results': [], 'errors': []}
    for index, submitted in enumerate(settings.get('files', []), 1):
        try:
            filename = _clean_name(submitted['name'])
            source = WORK / 'uploads' / str(index) / filename
            if source.suffix.lower() != '.xlsx':
                raise core.YoklamaHatasi('Yalnızca .xlsx desteklenmektedir')
            if not source.is_file():
                raise core.YoklamaHatasi('Yüklenen dosya bulunamadı')
            target, count, blanks, department, pages, warnings = core.yoklama_olustur(source, template, course, instructor)
            workbook = load_workbook(target, read_only=False)
            for sheet in workbook.worksheets:
                sheet['A1'] = title
                # Preserve heading styling while keeping the period title printable.
                from copy import copy
                heading_font = copy(sheet['A1'].font)
                heading_font.sz = min(heading_font.sz or 28, 24)
                sheet['A1'].font = heading_font
            workbook.save(target)
            sheet_data = [_sheet_data(sheet, imza, capacity) for sheet in workbook.worksheets]
            workbook.close()
            manifest['results'].append({
                'input': submitted['name'],
                'file': str(target),
                'filename': target.name,
                'count': count, 'blankRows': blanks, 'department': department,
                'pagesPerSheet': pages, 'warnings': warnings, 'sheets': sheet_data
            })
        except Exception as exc:
            manifest['errors'].append({'file': submitted.get('name', '(dosya)'), 'message': str(exc)})
    return json.dumps(manifest, ensure_ascii=False)
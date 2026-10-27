from copy import deepcopy
from pathlib import Path
from io import BytesIO
import re

import streamlit as st
from openpyxl import load_workbook
from pptx import Presentation

st.set_page_config(page_title="쿠팡 파렛트 적재리스트 생성기", page_icon="📦", layout="centered")

REQUIRED_HEADERS = ['파렛트', 'SKU No', 'SKU NAME', 'box수량', '수량', '제조일자']
TEMPLATE_PATH = Path(__file__).parent / '쿠팡 파렛트 양식.pptx'


def norm(v):
    return re.sub(r'\s+', '', str(v or '')).lower()


def cell_text(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_excel(file_obj):
    wb = load_workbook(file_obj, data_only=True, read_only=True)
    ws = wb.active
    meta = {}
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10), values_only=True):
        if len(row) >= 2 and row[0] is not None:
            key = norm(row[0])
            if key in ('입고예정일자', '납품센터명', '발주센터명', '발주번호'):
                meta[key] = cell_text(row[1])
    header_row = None
    headers = []
    for ridx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        normalized = [norm(x) for x in row]
        if all(norm(h) in normalized for h in REQUIRED_HEADERS):
            header_row, headers = ridx, list(row)
            break
    if not header_row:
        raise ValueError('엑셀에서 상품 표의 열 제목을 찾지 못했습니다. 필요한 열: ' + ', '.join(REQUIRED_HEADERS))
    idx = {norm(h): i for i, h in enumerate(headers) if h is not None}
    records = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(v is not None for v in row):
            continue
        pallet = cell_text(row[idx[norm('파렛트')]] if idx[norm('파렛트')] < len(row) else None)
        if not pallet:
            continue
        rec = {h: (row[idx[norm(h)]] if idx[norm(h)] < len(row) else None) for h in REQUIRED_HEADERS}
        rec['파렛트'] = pallet
        records.append(rec)
    if not records:
        raise ValueError('파렛트가 지정된 상품 행이 없습니다.')
    missing = [k for k in ('입고예정일자','납품센터명','발주센터명','발주번호') if not meta.get(k)]
    if missing:
        raise ValueError('기본 정보가 비어 있습니다: ' + ', '.join(missing))
    return meta, records


def replace_text(shape, new_text):
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    if not tf.paragraphs:
        tf.text = new_text
        return
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = new_text
        for run in p.runs[1:]:
            run.text = ''
    else:
        p.text = new_text
    for para in tf.paragraphs[1:]:
        for run in para.runs:
            run.text = ''


def set_cell(cell, value):
    text = cell_text(value)
    tf = cell.text_frame
    if not tf.paragraphs:
        cell.text = text
        return
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ''
    else:
        p.text = text
    for para in tf.paragraphs[1:]:
        for r in para.runs:
            r.text = ''


def clone_slide(prs, source_slide):
    slide = prs.slides.add_slide(source_slide.slide_layout)
    for shp in list(slide.shapes):
        shp._element.getparent().remove(shp._element)
    for shp in source_slide.shapes:
        newel = deepcopy(shp.element)
        slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
    for rel in source_slide.part.rels.values():
        if 'notesSlide' not in rel.reltype and 'slideLayout' not in rel.reltype:
            try:
                slide.part.rels.add_relationship(rel.reltype, rel._target, rel.rId, is_external=rel.is_external)
            except Exception:
                pass
    return slide


def rebuild_table(table, records):
    while len(table.rows) > 2:
        table._tbl.remove(table.rows[len(table.rows) - 1]._tr)
    template_tr = deepcopy(table.rows[1]._tr)
    if records:
        for i, rec in enumerate(records):
            if i == 0:
                row = table.rows[1]
            else:
                new_tr = table._tbl.add_tr(height=template_tr.h)
                for tc in list(new_tr.tc_lst):
                    new_tr.remove(tc)
                for tc in template_tr.tc_lst:
                    new_tr.append(deepcopy(tc))
                row = table.rows[len(table.rows) - 1]
            vals = [i + 1, rec['SKU No'], rec['SKU NAME'], rec['box수량'], rec['수량'], rec['제조일자']]
            for c, val in zip(row.cells, vals):
                set_cell(c, val)
    else:
        for c, val in zip(table.rows[1].cells, ['', '', '자료 없음', '', '', '']):
            set_cell(c, val)


def make_ppt(template_path, meta, records):
    prs = Presentation(template_path)
    if not prs.slides:
        raise ValueError('PPT 템플릿에 슬라이드가 없습니다.')
    first = prs.slides[0]
    slide_id_list = prs.slides._sldIdLst
    for sid in list(slide_id_list)[1:]:
        prs.part.drop_rel(sid.rId)
        slide_id_list.remove(sid)
    grouped = {}
    for rec in records:
        grouped.setdefault(rec['파렛트'], []).append(rec)
    keys = sorted(grouped, key=lambda x: (0, int(x)) if str(x).isdigit() else (1, str(x)))
    total = len(keys)
    for n, key in enumerate(keys, start=1):
        slide = first if n == 1 else clone_slide(prs, first)
        box_total = 0.0
        for r in grouped[key]:
            try:
                box_total += float(str(r['box수량'] or 0).replace(',', '').strip() or 0)
            except (TypeError, ValueError):
                raise ValueError(f'파렛트 {key}의 BOX 수량이 숫자가 아닙니다: {r["box수량"]}')
        box_label = str(int(box_total)) if box_total.is_integer() else str(box_total)
        for shape in slide.shapes:
            if shape.has_text_frame:
                if shape.text.startswith('  6 - 1') or ('총 박스수량' in shape.text and 'BOX' in shape.text):
                    replace_text(shape, f'  {total} - {n}                     / 총 박스수량  (   {box_label}   BOX)')
                elif '입고예정일자' in shape.text and '발주번호' in shape.text:
                    text = (f'입고예정일자 (   {meta["입고예정일자"]}  )   / 납품센터명 (      {meta["납품센터명"]}  )\n'
                            f'업체명    ( 주식회사 봉우리   )\t  / 발주센터명 (\t {meta["발주센터명"]} )\n'
                            f'발주번호      ({meta["발주번호"]} )')
                    replace_text(shape, text)
            if shape.has_table:
                rebuild_table(shape.table, grouped[key])
    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out, total, len(records)


st.title('📦 쿠팡 파렛트 적재리스트 생성기')
st.write('발주 엑셀을 업로드하면 파렛트별 PPT 적재리스트를 자동으로 만들어 드립니다.')
st.info('업체명은 “주식회사 봉우리”로 자동 입력되며, 파렛트별 BOX 합계와 표의 NO 번호를 자동 계산합니다.')

excel_file = st.file_uploader('① 발주 엑셀 파일 선택 (.xlsx / .xlsm)', type=['xlsx', 'xlsm'])

if not TEMPLATE_PATH.exists():
    st.error('PPT 양식 파일을 찾을 수 없습니다. GitHub 저장소에 “쿠팡 파렛트 양식.pptx” 파일이 있는지 확인해 주세요.')
elif excel_file is not None:
    if st.button('PPT 적재리스트 만들기', type='primary', use_container_width=True):
        try:
            with st.spinner('엑셀을 분석하고 PPT를 생성하는 중입니다...'):
                meta, records = read_excel(excel_file)
                result, pallet_count, row_count = make_ppt(TEMPLATE_PATH, meta, records)
            st.success(f'완료! 파렛트 {pallet_count}개, 상품 행 {row_count}개를 처리했습니다.')
            st.download_button(
                label='📥 완성된 PPT 다운로드',
                data=result.getvalue(),
                file_name=f'쿠팡_파렛트적재리스트_{re.sub(r"[^0-9가-힣]", "", meta["입고예정일자"]) or "입고예정"}.pptx',
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                use_container_width=True,
            )
            with st.expander('읽어온 기본 정보 확인'):
                st.write({
                    '입고예정일자': meta['입고예정일자'],
                    '납품센터명': meta['납품센터명'],
                    '발주센터명': meta['발주센터명'],
                    '발주번호': meta['발주번호'],
                    '업체명': '주식회사 봉우리',
                })
        except Exception as exc:
            st.error(f'파일을 생성하지 못했습니다: {exc}')
            st.caption('엑셀의 기본 정보와 표 제목(파렛트, SKU No, SKU NAME, box수량, 수량, 제조일자)이 있는지 확인해 주세요.')
else:
    st.caption('엑셀 파일을 선택하면 PPT 생성 버튼이 나타납니다.')

st.divider()
st.caption('주의: 실제 납품 전에 생성된 PPT의 품목, 수량, 박스 합계와 표 넘침 여부를 확인해 주세요.')

from copy import deepcopy
from datetime import date, datetime
from pathlib import Path
from io import BytesIO
import json
import re
import time
from urllib.request import Request, urlopen

import streamlit as st
from openpyxl import load_workbook
from pptx import Presentation

st.set_page_config(page_title="쿠팡 파렛트 적재리스트 생성기", page_icon="📦", layout="wide")

REQUIRED_HEADERS = ['파렛트', 'SKU No', 'SKU NAME', 'box수량', '수량', '제조일자']
TEMPLATE_PATH = Path(__file__).parent / '쿠팡 파렛트 양식.pptx'
PRODUCTS_PATH = Path(__file__).parent / 'products.json'
GITHUB_PRODUCTS_URL = 'https://raw.githubusercontent.com/ccfeel2-beep/coupang/main/products.json'


def norm(v):
    return re.sub(r'\s+', '', str(v or '')).lower()


def cell_text(v):
    if v is None:
        return ''
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or '').strip()
    for fmt in ('%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%Y.%m.%d', '%Y/%m/%d', '%Y%m%d', '%Y년 %m월 %d일'):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f'날짜 형식을 확인할 수 없습니다: {text}')

def korean_date(value):
    d = parse_date(value)
    return f'{d.month}월 {d.day}일'

def dotted_date(value):
    d = parse_date(value)
    return d.strftime('%Y.%m.%d')


def load_products_from_github():
    url = f"{GITHUB_PRODUCTS_URL}?v={int(time.time())}"
    request = Request(url, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache',
    })
    with urlopen(request, timeout=15) as response:
        raw_bytes = response.read()
    raw = json.loads(raw_bytes.decode('utf-8-sig'))
    return validate_products(raw)

def load_products_file():
    if not PRODUCTS_PATH.exists():
        return []
    with open(PRODUCTS_PATH, 'r', encoding='utf-8-sig') as f:
        raw = json.load(f)
    return validate_products(raw)



def initialize_products():
    try:
        return load_products_from_github(), 'GitHub', None
    except Exception as github_error:
        try:
            fallback = load_products_file()
            return fallback, '앱 내 예비 JSON', f'GitHub 상품 목록을 읽지 못했습니다: {github_error}'
        except Exception as local_error:
            return [], '불러오기 실패', (
                f'GitHub 상품 목록과 앱 내 예비 JSON을 모두 읽지 못했습니다. '
                f'GitHub 오류: {github_error} / 예비 JSON 오류: {local_error}'
            )


def validate_products(raw):
    if isinstance(raw, dict):
        raw = raw.get('products', [])
    if not isinstance(raw, list):
        raise ValueError('JSON은 products 배열을 포함해야 합니다.')
    result, seen = [], set()
    for i, item in enumerate(raw, start=1):
        if not isinstance(item, dict):
            raise ValueError(f'{i}번째 상품 형식이 올바르지 않습니다.')
        sku = str(item.get('sku_no', '')).strip()
        name = str(item.get('sku_name', '')).strip()
        active = item.get('active', True)
        if not sku or not name:
            raise ValueError(f'{i}번째 상품의 SKU No 또는 SKU Name이 비어 있습니다.')
        if sku in seen:
            raise ValueError(f'SKU No 중복: {sku}')
        if not isinstance(active, bool):
            raise ValueError(f'{sku}의 active 값은 true 또는 false여야 합니다.')
        seen.add(sku)
        result.append({'sku_no': sku, 'sku_name': name, 'active': active})
    return result


def read_excel(file_obj):
    wb = load_workbook(file_obj, data_only=True, read_only=True)
    ws = wb.active
    meta = {}
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10), values_only=True):
        if len(row) >= 2 and row[0] is not None:
            key = norm(row[0])
            if key in ('입고예정일자', '납품센터명', '발주센터명', '발주번호'):
                meta[key] = cell_text(row[1])
    header_row, headers = None, []
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
        grouped.setdefault(str(rec['파렛트']), []).append(rec)
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
                    text = (f'입고예정일자 (   {korean_date(meta["입고예정일자"])}  )   / 납품센터명 (      {meta["납품센터명"]}  )\n'
                            f'업체명    ( 주식회사 봉우리   )\t  발주센터명 (\t {meta["발주센터명"]} )\n'
                            f'발주번호      ({meta["발주번호"]} )')
                    replace_text(shape, text)
            if shape.has_table:
                rebuild_table(shape.table, [dict(r, 제조일자=dotted_date(meta['입고예정일자'])) for r in grouped[key]])
    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out, total, len(records)


def validate_job(meta, records):
    missing = [k for k in ('입고예정일자', '납품센터명', '발주센터명', '발주번호') if not str(meta.get(k, '')).strip()]
    if missing:
        raise ValueError('기본 정보가 비어 있습니다: ' + ', '.join(missing))
    if not records:
        raise ValueError('입력된 상품이 없습니다.')
    for i, rec in enumerate(records, start=1):
        if not str(rec.get('SKU No', '')).strip() or not str(rec.get('SKU NAME', '')).strip():
            raise ValueError(f'{i}번째 상품의 SKU 정보가 비어 있습니다.')
        for field in ('box수량', '수량'):
            try:
                if float(str(rec.get(field, 0)).replace(',', '')) < 0:
                    raise ValueError
            except Exception:
                raise ValueError(f'{i}번째 상품의 {field}은 0 이상의 숫자여야 합니다.')


if 'products' not in st.session_state:
    initial_products, initial_source, initial_error = initialize_products()
    st.session_state.products = initial_products
    st.session_state.products_source = initial_source
    st.session_state.products_load_error = initial_error
if 'pallet_items' not in st.session_state:
    st.session_state.pallet_items = {}
if 'next_item_id' not in st.session_state:
    st.session_state.next_item_id = 1

st.title('📦 쿠팡 파렛트 적재리스트 생성기')
st.write('엑셀 업로드 또는 웹 직접 입력으로 파렛트 적재리스트 PPT를 만듭니다.')
st.info('업체명은 주식회사 봉우리로 자동 입력됩니다. 상품 목록은 GitHub 저장소의 products.json에서 자동으로 불러옵니다.')

if not TEMPLATE_PATH.exists():
    st.error('PPT 양식 파일을 찾을 수 없습니다. GitHub 저장소에 “쿠팡 파렛트 양식.pptx” 파일이 있는지 확인해 주세요.')
else:
    tab_manual, tab_excel, tab_products = st.tabs(['✍️ 직접 입력', '📄 엑셀 업로드', '🗂️ 상품 관리'])

    with tab_products:
        st.subheader('상품 목록 관리')
        st.caption('상품 목록 원본은 GitHub 저장소의 products.json입니다. 사이트 접속 시 자동으로 불러오며, 중지 상품은 파렛트 입력 드롭다운에 표시되지 않습니다.')
        st.caption('GitHub 원본: ccfeel2-beep/coupang · main / products.json')

        if st.session_state.get('products_load_error'):
            st.warning(st.session_state.products_load_error)
        else:
            st.success(f"상품 목록을 {st.session_state.get('products_source', 'GitHub')}에서 불러왔습니다.")

        if st.button('🔄 GitHub 상품 목록 새로고침', use_container_width=True):
            try:
                latest_products = load_products_from_github()
                st.session_state.products = latest_products
                st.session_state.products_source = 'GitHub'
                st.session_state.products_load_error = None
                st.success(f'GitHub에서 상품 {len(latest_products)}개를 불러왔습니다.')
                st.rerun()
            except Exception as exc:
                st.error(f'GitHub 상품 목록을 새로고침하지 못했습니다. 현재 화면의 상품 목록은 유지됩니다. 오류: {exc}')

        with st.form('add_product_form', clear_on_submit=True):
            c1, c2 = st.columns(2)
            new_sku = c1.text_input('SKU No', placeholder='예: 10001')
            new_name = c2.text_input('SKU Name', placeholder='예: 태국 망고 5kg')
            add_clicked = st.form_submit_button('상품 등록', type='primary', use_container_width=True)
        if add_clicked:
            sku, name = new_sku.strip(), new_name.strip()
            if not sku or not name:
                st.error('SKU No와 SKU Name을 모두 입력해 주세요.')
            elif any(p['sku_no'] == sku for p in st.session_state.products):
                st.error('이미 등록된 SKU No입니다.')
            else:
                st.session_state.products.append({'sku_no': sku, 'sku_name': name, 'active': True})
                st.success(f'{sku} 상품을 등록했습니다. GitHub에 반영하려면 아래 JSON을 다운로드해 저장소의 products.json을 교체해 주세요.')
                st.rerun()

        if st.session_state.products:
            st.markdown('**등록된 상품**')
            for idx, product in enumerate(st.session_state.products):
                with st.container(border=True):
                    c1, c2, c3 = st.columns([1.1, 2.2, 1.2])
                    c1.write(product['sku_no'])
                    c2.write(product['sku_name'])
                    status_options = ['사용 중', '중지']
                    selected = c3.selectbox('상태', status_options, index=0 if product['active'] else 1, key=f"product_status_{product['sku_no']}", label_visibility='collapsed')
                    new_active = selected == '사용 중'
                    if new_active != product['active']:
                        st.session_state.products[idx]['active'] = new_active
                        st.rerun()
        else:
            st.warning('등록된 상품이 없습니다. 위에서 상품을 먼저 추가해 주세요.')

        st.divider()
        st.markdown('**GitHub에 저장할 상품 목록 다운로드**')
        json_data = json.dumps({'products': st.session_state.products}, ensure_ascii=False, indent=2).encode('utf-8')
        st.download_button('⬇️ 현재 상품 목록 JSON 다운로드', data=json_data, file_name='products.json', mime='application/json', use_container_width=True)
        st.warning('상품을 추가하거나 사용/중지 상태를 변경한 뒤 JSON을 다운로드하고, GitHub 저장소의 최상위 products.json 파일을 교체해 Commit changes를 눌러 주세요. GitHub에 반영된 상품 목록은 다음 사이트 접속 때 자동으로 불러옵니다. 새로고침은 GitHub 원본으로 현재 화면의 변경사항을 덮어쓰므로, 저장하지 않은 변경사항이 있다면 먼저 JSON을 다운로드하세요.')

    with tab_manual:
        st.subheader('납품 기본정보')
        c1, c2 = st.columns(2)
        delivery_date = c1.date_input('입고예정일자', value=date.today())
        order_no = c2.text_input('발주번호')

        # 세션 상태에 기본 센터명이 없으면 '인천17'로 초기화
        if 'delivery_center' not in st.session_state:
            st.session_state.delivery_center = '인천17'
        if 'order_center' not in st.session_state:
            st.session_state.order_center = '인천17'

        c3, c4 = st.columns(2)
        # key 매개변수를 지정하면 사용자가 바꾼 값이 st.session_state에 자동으로 저장되고 고정됩니다.
        delivery_center = c3.text_input('납품센터명', key='delivery_center')
        order_center = c4.text_input('발주센터명', key='order_center')

        st.caption('업체명: 주식회사 봉우리 (자동 입력) · 제조일자는 입고예정일자와 동일하게 자동 적용됩니다.')
        st.divider()
        st.subheader('파렛트별 상품 배치')
        pallet_count = st.number_input('파렛트 개수', min_value=1, max_value=50, value=6, step=1)
        # 수정: active_products를 sku_no 기준(오름차순)으로 정렬
        active_products = [p for p in st.session_state.products if p['active']]
        if not active_products:
            st.warning('사용 중인 상품이 없습니다. 상품 관리 메뉴에서 상품을 등록하거나 사용 중으로 변경해 주세요.')
        else:
            # sku_no를 기준으로 정렬 (숫자/문자 오름차순)
            sorted_active_products = sorted(active_products, key=lambda p: str(p['sku_no']))
            product_labels = {f"{p['sku_no']} | {p['sku_name']}": p for p in sorted_active_products}  
            labels = list(product_labels.keys())
            manual_records = []
            for pallet_no in range(1, int(pallet_count) + 1):
                with st.expander(f'파렛트 {pallet_no}', expanded=(pallet_no == 1)):
                    if pallet_no not in st.session_state.pallet_items:
                        st.session_state.pallet_items[pallet_no] = [st.session_state.next_item_id]
                        st.session_state.next_item_id += 1
                    item_ids = st.session_state.pallet_items[pallet_no]
                    st.caption('상품                                      BOX 수량                 수량')
                    for item_id in list(item_ids):
                        a, b, c, d = st.columns([3.5, 1, 1, 0.35])
                        selected_label = a.selectbox('상품', labels, key=f'manual_product_{pallet_no}_{item_id}', label_visibility='collapsed')
                        box_qty = b.number_input('BOX 수량', min_value=0, step=1, key=f'manual_box_{pallet_no}_{item_id}', label_visibility='collapsed')
                        qty = c.number_input('수량', min_value=0, step=1, key=f'manual_qty_{pallet_no}_{item_id}', label_visibility='collapsed')
                        if d.button('✕', key=f'delete_item_{pallet_no}_{item_id}', help='이 상품 행 삭제', use_container_width=True):
                            st.session_state.pallet_items[pallet_no].remove(item_id)
                            st.rerun()
                        chosen = product_labels[selected_label]
                        manual_records.append({'파렛트': str(pallet_no), 'SKU No': chosen['sku_no'], 'SKU NAME': chosen['sku_name'], 'box수량': box_qty, '수량': qty, '제조일자': dotted_date(delivery_date)})
                    if st.button('＋ 상품 추가', key=f'add_row_{pallet_no}', use_container_width=True):
                        st.session_state.pallet_items[pallet_no].append(st.session_state.next_item_id)
                        st.session_state.next_item_id += 1
                        st.rerun()
                    if not st.session_state.pallet_items[pallet_no]:
                        st.caption('등록된 상품 행이 없습니다. 상품 추가를 눌러주세요.')
            st.divider()
            st.subheader('PPT 생성')
            if st.button('📦 직접 입력한 내용으로 PPT 만들기', type='primary', use_container_width=True):
                try:
                    meta = {'입고예정일자': delivery_date.strftime('%Y-%m-%d'), '납품센터명': delivery_center.strip(), '발주센터명': order_center.strip(), '발주번호': order_no.strip()}
                    validate_job(meta, manual_records)
                    with st.spinner('PPT를 생성하고 있습니다...'):
                        result, pallet_total, row_total = make_ppt(TEMPLATE_PATH, meta, manual_records)
                    st.success(f'완료! 파렛트 {pallet_total}개, 상품 행 {row_total}개를 처리했습니다.')
                    st.download_button('📥 완성된 PPT 다운로드', data=result.getvalue(), file_name=f'쿠팡_파렛트적재리스트_{delivery_date.strftime("%Y%m%d")}.pptx', mime='application/vnd.openxmlformats-officedocument.presentationml.presentation', use_container_width=True, key='manual_ppt_download')
                except Exception as exc:
                    st.error(f'PPT를 생성하지 못했습니다: {exc}')

    with tab_excel:
        st.subheader('기존 엑셀 업로드 방식')
        st.write('기존 발주 엑셀을 업로드하면 파렛트별 PPT 적재리스트를 자동으로 만듭니다.')
        excel_file = st.file_uploader('발주 엑셀 파일 선택 (.xlsx / .xlsm)', type=['xlsx', 'xlsm'], key='excel_upload')
        if excel_file is not None and st.button('엑셀로 PPT 적재리스트 만들기', type='primary', use_container_width=True):
            try:
                with st.spinner('엑셀을 분석하고 PPT를 생성하는 중입니다...'):
                    meta, records = read_excel(excel_file)
                    result, pallet_total, row_total = make_ppt(TEMPLATE_PATH, meta, records)
                st.success(f'완료! 파렛트 {pallet_total}개, 상품 행 {row_total}개를 처리했습니다.')
                st.download_button('📥 완성된 PPT 다운로드', data=result.getvalue(), file_name=f'쿠팡_파렛트적재리스트_{re.sub(r"[^0-9가-힣]", "", meta["입고예정일자"]) or "입고예정"}.pptx', mime='application/vnd.openxmlformats-officedocument.presentationml.presentation', use_container_width=True, key='excel_ppt_download')
                with st.expander('읽어온 기본 정보 확인'):
                    st.write({**meta, '업체명': '주식회사 봉우리'})
            except Exception as exc:
                st.error(f'파일을 생성하지 못했습니다: {exc}')

st.divider()
st.caption('주의: 실제 납품 전에 생성된 PPT의 품목, 수량, 박스 합계와 표 넘침 여부를 확인해 주세요.')

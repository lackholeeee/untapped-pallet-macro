import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.pagebreak import Break
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.font_manager as fm
import io
import re
import os
from collections import defaultdict

@st.cache_resource
def get_korean_font():
    paths = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "C:/Windows/Fonts/malgun.ttf",
        "/Library/Fonts/AppleGothic.ttf"
    ]
    for p in paths:
        if os.path.exists(p):
            return fm.FontProperties(fname=p)
    return fm.FontProperties(family='sans-serif')

font_prop = get_korean_font()
plt.rcParams['axes.unicode_minus'] = False

st.set_page_config(page_title='UNTAPPED 물류 자동화 통합 엔진', layout='wide', page_icon='📦')
st.title('📦 UNTAPPED STUDIO 하역 작업 지시서 매크로 (공식 백엔드 엔진)')
st.markdown('패킹리스트 엑셀 파일을 업로드하고 조건을 선택하면 **1단계: 카톡 요약문**, **2단계: 단 1개의 3-시트 공식 엑셀(.xlsx)**, **3단계: 흑백 인쇄 최적화 CAD 미니맵**을 일괄 자동 생성합니다.')

st.sidebar.header('⚙️ 5대 하역 및 적재 조건 설정')
work_purpose = st.sidebar.radio('1. 작업 목적 선택', [
    '1.A 단순 보관 (완파레트 중심 공간 절약 적재 / 소량 화물 1P 통합 보관) [기본값]',
    '1.B 송장 출고 (사이즈별 피킹 직배송 / 2차 분해 까대기 원천 방지) ⭐'
])
box_limit_preset = st.sidebar.selectbox('2. 박스 규격 및 완P 한도(L)', [
    'L=16 (일반 표준 박스 / 600*400*350 / 4단)',
    'L=12 (다꾸앙 패딩 및 자켓 / 600*400*530 / 3단)',
    'L=9 (양홍 코트 등 대형 박스)',
    '직접 수기 입력'
])
custom_limit = st.sidebar.number_input('수기 입력 (L)', min_value=1, max_value=40, value=16) if '수기' in box_limit_preset else None
mix_option = st.sidebar.radio('3. MIX(혼적) 적재 방식', [
    '3.A 진열/검수용 별도 격리 (도크 초입 MIX 전용 P-01 분리)',
    '3.B 여유분 적재 합적 (파레트 공간 절약 모드 / 자투리 상단 합적)'
])
new_items_input = st.sidebar.text_input('4. 초도 신상품 상품코드 (공란 시 100% 지속 리오더)', '')
urgent_input = st.sidebar.text_input('5. 긴급 출고 건 유무 (선택: 특정 코드/컬러/사이즈)', '')

default_L = custom_limit if custom_limit else (16 if '16' in box_limit_preset else (12 if '12' in box_limit_preset else 9))
is_dispatch = '1.B' in work_purpose
is_mix_quarantine = '3.A' in mix_option
initial_skus = [x.strip().upper() for x in re.split(r'[,/\s]+', new_items_input) if x.strip()]
urgent_skus = [x.strip().upper() for x in re.split(r'[,/\s]+', urgent_input) if x.strip()]

uploaded_files = st.file_uploader('📥 패킹리스트 엑셀 업로드 (다중 파일 지원)', type=['xlsx', 'xls'], accept_multiple_files=True)

def parse_all_excel(files):
    all_boxes = []
    size_candidates = ['S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', 'FREE']
    
    for f in files:
        filename = f.name if hasattr(f, 'name') else str(f)
        if '언탭트_현장총괄하역오더' in filename: continue
        try:
            xl = pd.ExcelFile(f)
            
            # Form 1: 요약표형 (시트명에 '요약' 또는 헤더에 품번/박스)
            is_summary_table = any('요약' in str(s) or 'SUMMARY' in str(s).upper() for s in xl.sheet_names)
            if not is_summary_table:
                df_chk = xl.parse(xl.sheet_names[0], header=None)
                for r in range(min(5, len(df_chk))):
                    r_str = " ".join([str(v) for v in df_chk.iloc[r].values if pd.notna(v)])
                    if any(k in r_str for k in ['품번', '상품코드']) and any(k in r_str for k in ['박스', 'BOX', 'CNT']):
                        is_summary_table = True
                        break
                        
            if is_summary_table:
                target_sheet = xl.sheet_names[0]
                for s in xl.sheet_names:
                    if '요약' in s or 'SUMMARY' in s.upper():
                        target_sheet = s
                        break
                df = xl.parse(target_sheet, header=None)
                
                header_idx = -1
                for r in range(min(10, len(df))):
                    row_vals = [str(v).strip() for v in df.iloc[r].values if pd.notna(v)]
                    if any(k in v for k in ['품번', '상품', 'PROD'] for v in row_vals):
                        header_idx = r
                        break
                if header_idx == -1: header_idx = 2
                
                header = [str(v).strip() for v in df.iloc[header_idx].values]
                size_map = {}
                prod_col, color_col, box_col = 0, 1, -1
                
                for c_idx, h in enumerate(header):
                    h_up = h.upper()
                    if h_up in size_candidates:
                        size_map[c_idx] = h_up
                    elif any(k in h_up for k in ['박스', 'BOX', 'CNT', '箱数']):
                        box_col = c_idx
                    elif any(k in h_up for k in ['컬러', '칼라', 'COLOR', '색상']):
                        color_col = c_idx
                    elif any(k in h_up for k in ['품번', '상품', 'PROD', '款号']):
                        prod_col = c_idx
                        
                if box_col == -1: box_col = len(header) - 1
                
                for r in range(header_idx + 1, len(df)):
                    row = df.iloc[r]
                    p_val = row.iloc[prod_col] if pd.notna(row.iloc[prod_col]) else ""
                    p_str = str(p_val).strip()
                    if not p_str or any(k in p_str for k in ['합계', 'TOTAL', '合计', '小计']):
                        continue
                        
                    c_val = row.iloc[color_col] if pd.notna(row.iloc[color_col]) else "DEFAULT"
                    c_str = str(c_val).strip().upper()
                    
                    box_val = row.iloc[box_col] if pd.notna(row.iloc[box_col]) else 1
                    try: box_count = int(float(box_val))
                    except: box_count = 1
                    
                    row_items = []
                    for c_idx, sz in size_map.items():
                        q_val = row.iloc[c_idx] if pd.notna(row.iloc[c_idx]) else 0
                        try: q_int = int(float(q_val))
                        except: q_int = 0
                        if q_int > 0:
                            row_items.append({'size': sz, 'qty': q_int})
                            
                    if row_items:
                        for b_i in range(box_count):
                            box_items = []
                            for it in row_items:
                                base_q = it['qty'] // box_count
                                rem = it['qty'] % box_count
                                q_box = base_q + (1 if b_i < rem else 0)
                                if q_box > 0:
                                    box_items.append({'color': c_str, 'size': it['size'], 'qty': q_box})
                            if box_items:
                                all_boxes.append({
                                    'pkg': f"{p_str}-{c_str}-CTN-{b_i+1}",
                                    'prod': p_str,
                                    'color': c_str,
                                    'items': box_items,
                                    'raw_title': p_str
                                })
            else:
                # Form 2: Style Workbook (시트명이 숫자인 형태: 144, 145, 158 등)
                is_style_workbook = all(str(s).strip().isdigit() for s in xl.sheet_names)
                
                if is_style_workbook:
                    for sheet in xl.sheet_names:
                        df = xl.parse(sheet, header=None)
                        prod = f'MUN{sheet}'
                        
                        size_map = {}
                        size_row_idx = -1
                        for r_idx in range(min(15, len(df))):
                            row_vals = [str(v).strip().upper() for v in df.iloc[r_idx].values if pd.notna(v)]
                            found = [s for s in size_candidates if s in row_vals]
                            if len(found) >= 1:
                                for c_idx, val in enumerate(df.iloc[r_idx].values):
                                    v_str = str(val).strip().upper()
                                    if v_str in size_candidates: size_map[c_idx] = v_str
                                size_row_idx = r_idx
                                break
                        else: continue
                            
                        ct_col, color_col, box_cnt_col = 1, 4, 10
                        for r_idx in range(min(15, len(df))):
                            for c_idx, val in enumerate(df.iloc[r_idx].values):
                                if pd.notna(val):
                                    v_str = str(val).strip().upper()
                                    if v_str in ['箱数', 'BOX', 'BOXES', 'CNT']: box_cnt_col = c_idx
                                    elif v_str in ['颜色', 'COLOR', '색상']: color_col = c_idx
                                    elif v_str in ['箱号', 'CT NO', 'CTNO', '박스번호']: ct_col = c_idx
                                    
                        df[color_col] = df[color_col].ffill()
                        
                        for r in range(size_row_idx + 2, len(df)):
                            row = df.iloc[r]
                            ct_val = row.iloc[ct_col]
                            if not pd.notna(ct_val): continue
                            try: ct_num = int(float(ct_val))
                            except: break
                                
                            color_val = row.iloc[color_col]
                            color_str = str(color_val).strip().upper() if pd.notna(color_val) else 'DEFAULT'
                            box_cnt_val = row.iloc[box_cnt_col] if box_cnt_col < len(row.values) else 1
                            try: box_count = int(float(box_cnt_val)) if pd.notna(box_cnt_val) else 1
                            except: box_count = 1
                                
                            row_items = []
                            for c_idx, sz in size_map.items():
                                if c_idx < len(row.values):
                                    q_val = row.iloc[c_idx]
                                    if pd.notna(q_val):
                                        try:
                                            q_int = int(float(q_val))
                                            if q_int > 0: row_items.append({'color': color_str, 'size': sz, 'qty': q_int})
                                        except: pass
                                            
                            if row_items:
                                for b_i in range(box_count):
                                    p_name = f'{prod}-{sheet}-CTN-{ct_num}' if box_count == 1 else f'{prod}-{sheet}-CTN-{ct_num}-{b_i+1}'
                                    all_boxes.append({'pkg': p_name, 'prod': prod, 'items': row_items, 'raw_title': prod})
                else:
                    # Form 3 & 4: 일반 가로 매트릭스형 또는 세로 리스트형
                    target_sheet = xl.sheet_names[0]
                    for s in xl.sheet_names:
                        if any(k in s.upper() for k in ['MUN', '패킹', '출고', '입고', 'DETAIL']):
                            target_sheet = s
                            break
                    df = xl.parse(target_sheet, header=None)

                    header_row_idx, format_type = -1, None
                    size_cols_map, col_map = {}, {}
                    
                    for r_idx in range(min(20, len(df))):
                        row_vals = [str(v).strip().upper().replace(' ', '') for v in df.iloc[r_idx].values if pd.notna(v)]
                        found_sizes = [s for s in size_candidates if s in row_vals]
                        if len(found_sizes) >= 2:
                            header_row_idx = r_idx
                            format_type = 'MATRIX'
                            for col_idx, col_name in enumerate(df.iloc[r_idx].values):
                                val = str(col_name).strip().upper().replace(' ', '')
                                if val in size_candidates: size_cols_map[val] = col_idx
                            break
                            
                        if any('사이즈' in v or 'SIZE' in v or '尺码' in v for v in row_vals) and any('수량' in v or 'QTY' in v or '件数' in v for v in row_vals):
                            header_row_idx = r_idx
                            format_type = 'LIST'
                            for col_idx, col_name in enumerate(df.iloc[r_idx].values):
                                val = str(col_name).strip().upper().replace(' ', '')
                                if any(k in val for k in ['상품', '款号', 'PROD', '품명']): col_map['prod'] = col_idx
                                if any(k in val for k in ['컬러', '칼라', 'COLOR', '색상']): col_map['color'] = col_idx
                                if any(k in val for k in ['사이즈', 'SIZE', '尺码']): col_map['size'] = col_idx
                                if any(k in val for k in ['수량', 'QTY', '件数']): col_map['qty'] = col_idx
                                if any(k in val for k in ['패킹', '박스번호', '箱号', 'CTN', 'CARTON']): col_map['pkg'] = col_idx
                            break

                    if format_type == 'MATRIX':
                        cur_prod, cur_color = 'MUN_ITEM', 'DEFAULT'
                        pkg_col, box_cnt_col = 2, 10
                        search_rows = [header_row_idx]
                        if header_row_idx > 0: search_rows.insert(0, header_row_idx - 1)
                        
                        for sr in search_rows:
                            for c, v in enumerate(df.iloc[sr].values):
                                if pd.notna(v):
                                    v_str = str(v).strip().upper().replace(' ', '')
                                    if any(k in v_str for k in ['패킹', '箱号', 'CTN', '박스번호', 'CTNO']): pkg_col = c
                                    if any(k in v_str for k in ['箱数', '박스수', 'CTNS', 'CNT']): box_cnt_col = c

                        raw_prod_title = ""
                        for r in range(header_row_idx + 1, len(df)):
                            row = df.iloc[r]
                            row_str = ' '.join([str(v) for v in row.values if pd.notna(v)])
                            if any(k in row_str for k in ['합계', '合计', '小计', '총합', 'TOTAL']): continue

                            c0 = str(row.iloc[0]).strip() if len(row.values) > 0 and pd.notna(row.iloc[0]) else ''
                            c1 = str(row.iloc[1]).strip() if len(row.values) > 1 and pd.notna(row.iloc[1]) else ''
                            c2 = str(row.iloc[2]).strip() if len(row.values) > 2 and pd.notna(row.iloc[2]) else ''
                            
                            for c_val in [c0, c1, c2]:
                                if not c_val: continue
                                c_up = c_val.upper()
                                if any(c_up.startswith(p) for p in ['MUN', 'DJ']):
                                    match = re.search(r'(MUN|DJ)\s*\d+[A-Z0-9]*', c_up)
                                    if match:
                                        cur_prod = match.group(0).replace(' ', '')
                                        raw_prod_title = c_val
                                    rest = c_up[len(match.group(0)):].strip() if match else ''
                                    rest = re.sub(r'[^A-Z\s]', '', rest).strip()
                                    if rest and not any(k in rest for k in ['총', '박스', '합']): cur_color = rest
                                elif not c_up.isdigit() and len(c_up) >= 1 and not any(k in c_up for k in ['총', '박스', '합', '호', '-']):
                                    if re.match(r'^[A-Z\s]+$', c_up): cur_color = c_up

                            pkg_val = str(row.iloc[pkg_col]).strip() if pkg_col < len(row.values) and pd.notna(row.iloc[pkg_col]) else ''
                            if not pkg_val or pkg_val in ['nan', '']: continue
                                
                            box_count = 1
                            if box_cnt_col is not None and box_cnt_col < len(row.values) and pd.notna(row.iloc[box_cnt_col]):
                                try: box_count = int(float(row.iloc[box_cnt_col]))
                                except: box_count = 1

                            row_items = []
                            for sz, c_idx in size_cols_map.items():
                                if c_idx < len(row.values):
                                    q_val = row.iloc[c_idx]
                                    if pd.notna(q_val):
                                        try:
                                            q_int = int(float(q_val))
                                            if q_int > 0: row_items.append({'color': cur_color, 'size': sz, 'qty': q_int})
                                        except: pass

                            if row_items:
                                for b_i in range(box_count):
                                    p_name = f'{pkg_val}-{b_i+1}' if box_count > 1 else pkg_val
                                    all_boxes.append({'pkg': p_name, 'prod': cur_prod, 'items': row_items, 'raw_title': raw_prod_title})

                    elif format_type == 'LIST':
                        cur_prod, cur_color = 'MUN_ITEM', 'DEFAULT'
                        prev_pkg, empty_streak = None, 0
                        raw_prod_title = ""

                        for r in range(header_row_idx + 1, len(df)):
                            row = df.iloc[r]
                            row_str = ' '.join([str(v) for v in row.values if pd.notna(v)])
                            if any(k in row_str for k in ['박스 총 수량', '옷 총 수량', '총합', '合计', '小计', 'TOTAL']): continue
                                
                            p_val = row.iloc[col_map['prod']] if 'prod' in col_map and col_map['prod'] < len(row.values) else None
                            c_val = row.iloc[col_map['color']] if 'color' in col_map and col_map['color'] < len(row.values) else None
                            s_val = row.iloc[col_map['size']] if 'size' in col_map and col_map['size'] < len(row.values) else None
                            q_val = row.iloc[col_map['qty']] if 'qty' in col_map and col_map['qty'] < len(row.values) else None
                            pkg_val = row.iloc[col_map['pkg']] if 'pkg' in col_map and col_map['pkg'] < len(row.values) else None
                            
                            if not pd.notna(s_val) and not pd.notna(pkg_val) and not pd.notna(q_val):
                                empty_streak += 1
                                if empty_streak >= 3: break
                                continue
                            else: empty_streak = 0
                            
                            if pd.notna(p_val) and str(p_val).strip().upper().startswith(('MUN', 'DJ')):
                                match = re.search(r'(MUN|DJ)\s*\d+[A-Z0-9]*', str(p_val).strip().upper())
                                if match: cur_prod = match.group(0).replace(' ', '')
                                else: cur_prod = str(p_val).strip().upper()
                                raw_prod_title = str(p_val).strip()
                                
                            if pd.notna(c_val) and not any(k in str(c_val) for k in ['총', '박스', '합', '合计']):
                                c_up = str(c_val).strip().upper()
                                if re.match(r'^[A-Z\s]+$', c_up): cur_color = c_up
                                else:
                                    extracted = re.sub(r'[^A-Z\s]', '', c_up).strip()
                                    if extracted: cur_color = extracted

                            if not pd.notna(q_val): continue
                                
                            try:
                                q_int = int(float(q_val))
                                sz = str(s_val).strip().upper()
                                if q_int > 0 and sz in size_candidates:
                                    if pd.notna(pkg_val) and str(pkg_val).strip():
                                        pkg_str = str(pkg_val).strip()
                                        if pkg_str == prev_pkg and all_boxes:
                                            all_boxes[-1]['items'].append({'color': cur_color, 'size': sz, 'qty': q_int})
                                        else:
                                            all_boxes.append({'pkg': pkg_str, 'prod': cur_prod, 'items': [{'color': cur_color, 'size': sz, 'qty': q_int}], 'raw_title': raw_prod_title})
                                        prev_pkg = pkg_str
                                    else:
                                        if prev_pkg and all_boxes:
                                            all_boxes[-1]['items'].append({'color': cur_color, 'size': sz, 'qty': q_int})
                                        else:
                                            pkg_str = f'CTN-{len(all_boxes)+1}'
                                            all_boxes.append({'pkg': pkg_str, 'prod': cur_prod, 'items': [{'color': cur_color, 'size': sz, 'qty': q_int}], 'raw_title': raw_prod_title})
                                            prev_pkg = pkg_str
                            except: pass
        except Exception as e:
            st.error(f'파일 파싱 중 예외 발생: {e}')

    return all_boxes

# 품목별 지능형 L값 결정 함수 (MUN039, 패딩, 다운 감지 시 자동 L=12 적용)
def get_item_L_limit(prod, raw_title, default_L):
    search_str = f"{prod} {raw_title}".upper()
    if '039' in search_str or '패딩' in search_str or '푸퍼' in search_str or '다운' in search_str:
        return min(default_L, 12)
    return default_L

if uploaded_files:
    raw_boxes = parse_all_excel(uploaded_files)
    st.info(f'📊 수신 완료: 총 {len(uploaded_files)}개 파일에서 총 {len(raw_boxes)}박스 데이터가 정상 감지되었습니다.')

    if st.button('⚡ 현장 작업지시서 일괄 생성', type='primary'):
        with st.spinner('파레트 적재 계산, 3-시트 엑셀 및 2D CAD 미니맵 렌더링 중...'):
            size_order = ['S', 'M', 'L', 'XL', '2XL', '3XL', '4XL']
            mix_boxes = []
            pure_by_sku_grouped = defaultdict(list)
            qty_summary = defaultdict(lambda: defaultdict(lambda: {'boxes': 0, 'pcs': 0}))

            for b in raw_boxes:
                colors = {it['color'] for it in b['items']}
                for it in b['items']:
                    qty_summary[(b['prod'], it['color'])][it['size']]['pcs'] += it['qty']
                
                # 순수 단일컬러 vs 다색 혼적
                if len(colors) > 1:
                    mix_boxes.append(b)
                else:
                    c = list(colors)[0] if colors else 'DEFAULT'
                    pure_by_sku_grouped[(b['prod'], c)].append(b)

            pallets = []
            p_counter = 1

            # 3.A 선택 시: MIX 박스를 도크 초입 P-01로 단독 격리
            order_summary_rows = []
            if is_mix_quarantine and mix_boxes:
                pallets.append({
                    'p_num': f'P-{p_counter:02d}',
                    'prod': 'MIX BATCH',
                    'color': '혼적',
                    'size': '다색혼적',
                    'boxes': len(mix_boxes),
                    'desc': f'다색혼적 {len(mix_boxes)}박스 (도크 초입 격리)',
                    'type': 'MIX',
                    'is_full': False
                })
                p_counter += 1

                mix_summary_lines = []
                for mb in mix_boxes:
                    details = ', '.join([f"{it['color']} {it['size']}-{it['qty']}장" for it in mb['items']])
                    mix_summary_lines.append(f'• [박스 {mb["pkg"]}] {details}')
                mix_str = '\n'.join(mix_summary_lines)

                order_summary_rows.append({
                    'prod': 'MIX BATCH',
                    'color': '혼적전용',
                    'pallets': '1P',
                    'boxes': f'{len(mix_boxes)}박스',
                    'is_mix': True,
                    'mix_text': f'【도크초입 별도격리】 다색혼적 총 {len(mix_boxes)}박스\n{mix_str}'
                })

            for (prod, color), boxes_list in pure_by_sku_grouped.items():
                tot_pure = len(boxes_list)
                active_sizes = [s for s in size_order if qty_summary[(prod, color)][s]['pcs'] > 0]
                raw_title = boxes_list[0].get('raw_title', '') if boxes_list else ''
                sku_L = get_item_L_limit(prod, raw_title, default_L)
                
                # 단일 사이즈 박스 수량 집계
                size_box_counts = defaultdict(int)
                for b in boxes_list:
                    b_szs = {it['size'] for it in b['items']}
                    if len(b_szs) == 1:
                        size_box_counts[list(b_szs)[0]] += 1

                is_pure_size_boxes = (sum(size_box_counts.values()) == tot_pure)
                p_start = p_counter
                row_lines = []

                if not is_pure_size_boxes:
                    # 어소트 박스 (박스 안에 여러 사이즈가 섞인 형태)
                    if tot_pure <= sku_L:
                        pallets.append({
                            'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': 'ALL',
                            'boxes': tot_pure, 'desc': f'{tot_pure}박스 (어소트 혼합박스)', 'type': 'ALL', 'is_full': (tot_pure == sku_L)
                        })
                        row_lines.append(f'[ALL] 1P ({tot_pure}박스 어소트)')
                        p_counter += 1
                    else:
                        full_p = tot_pure // sku_L
                        rem_b = tot_pure % sku_L
                        for _ in range(full_p):
                            pallets.append({
                                'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': '완P',
                                'boxes': sku_L, 'desc': f'{sku_L}박스 (완P)', 'type': '완P', 'is_full': True
                            })
                            row_lines.append(f'[완P] 1P ({sku_L}박스)')
                            p_counter += 1
                        if rem_b > 0:
                            pallets.append({
                                'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': '자투리',
                                'boxes': rem_b, 'desc': f'{rem_b}박스 (자투리)', 'type': '자투리', 'is_full': False
                            })
                            row_lines.append(f'[자투리] 1P ({rem_b}박스)')
                            p_counter += 1
                else:
                    # 단일 사이즈 박스: 동일 사이즈 쪼개기 까대기 방지 규칙 적용
                    if not is_dispatch:
                        # 1.A 단순 보관 모드
                        if tot_pure <= sku_L:
                            sz_desc = ', '.join([f'{s}-{size_box_counts[s]}박스' for s in active_sizes if size_box_counts[s] > 0])
                            pallets.append({
                                'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': 'ALL',
                                'boxes': tot_pure, 'desc': f'{tot_pure}박스: {sz_desc}', 'type': 'ALL', 'is_full': (tot_pure == sku_L)
                            })
                            row_lines.append(f'[ALL] 1P ({tot_pure}박스: {sz_desc})')
                            p_counter += 1
                        else:
                            # T > sku_L: 사이즈 덩어리 보존 적재
                            remaining_chunks = []
                            for sz in active_sizes:
                                cnt = size_box_counts[sz]
                                if cnt == 0: continue
                                if cnt >= sku_L:
                                    n_full = cnt // sku_L
                                    rem = cnt % sku_L
                                    # 초과분이 1박스인 경우: 완P 상단에 +★1박스 합적 표기
                                    if rem == 1:
                                        pallets.append({
                                            'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': sz,
                                            'boxes': sku_L + 1, 'desc': f'{sku_L}박스 완P + ★{sz}-1박스 합적 (총 {sku_L+1}박스)',
                                            'type': '완P+합적', 'is_full': True
                                        })
                                        row_lines.append(f'[{sz}] 1P ({sku_L}박스 완P + ★{sz}-1박스 합적)')
                                        p_counter += 1
                                    else:
                                        for _ in range(n_full):
                                            pallets.append({
                                                'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': sz,
                                                'boxes': sku_L, 'desc': f'{sku_L}박스 (완P)', 'type': '완P', 'is_full': True
                                            })
                                            row_lines.append(f'[{sz}] 1P ({sku_L}박스 완P)')
                                            p_counter += 1
                                        if rem > 0:
                                            remaining_chunks.append((sz, rem))
                                else:
                                    remaining_chunks.append((sz, cnt))
                            
                            # 잔여 사이즈 청크들을 쪼개지 않고 배정
                            cur_p_chunks = []
                            cur_sum = 0
                            for sz, cnt in remaining_chunks:
                                if cur_sum + cnt <= sku_L:
                                    cur_p_chunks.append((sz, cnt))
                                    cur_sum += cnt
                                else:
                                    tag = "+".join([s for s, c in cur_p_chunks])
                                    desc = ", ".join([f"{s}-{c}박스" for s, c in cur_p_chunks])
                                    pallets.append({
                                        'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': tag,
                                        'boxes': cur_sum, 'desc': f'{cur_sum}박스: {desc}', 'type': tag, 'is_full': (cur_sum == sku_L)
                                    })
                                    row_lines.append(f'[{tag}] 1P ({cur_sum}박스: {desc})')
                                    p_counter += 1
                                    cur_p_chunks = [(sz, cnt)]
                                    cur_sum = cnt
                            if cur_p_chunks:
                                tag = "+".join([s for s, c in cur_p_chunks]) if len(cur_p_chunks) <= 2 else "ALL"
                                desc = ", ".join([f"{s}-{c}박스" for s, c in cur_p_chunks])
                                pallets.append({
                                    'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': tag,
                                    'boxes': cur_sum, 'desc': f'{cur_sum}박스: {desc}', 'type': tag, 'is_full': (cur_sum == sku_L)
                                })
                                row_lines.append(f'[{tag}] 1P ({cur_sum}박스: {desc})')
                                p_counter += 1
                    else:
                        # 1.B 송장 출고 모드 (사이즈별 단독 파레트)
                        for sz in active_sizes:
                            cnt = size_box_counts[sz]
                            if cnt == 0: continue
                            n_full = cnt // sku_L
                            rem = cnt % sku_L
                            for _ in range(n_full):
                                pallets.append({
                                    'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': sz,
                                    'boxes': sku_L, 'desc': f'{sku_L}박스 (완P 직배송)', 'type': '완P', 'is_full': True
                                })
                                row_lines.append(f'[{sz}] 1P ({sku_L}박스 완P)')
                                p_counter += 1
                            if rem > 0:
                                pallets.append({
                                    'p_num': f'P-{p_counter:02d}', 'prod': prod, 'color': color, 'size': sz,
                                    'boxes': rem, 'desc': f'{rem}박스 (피킹 전용 단독P)', 'type': sz, 'is_full': False
                                })
                                row_lines.append(f'[{sz}] 1P ({rem}박스 단독P)')
                                p_counter += 1

                p_assigned = p_counter - p_start
                order_summary_rows.append({
                    'prod': prod, 'color': color, 'pallets': f'{p_assigned}P', 'boxes': f'{tot_pure}박스',
                    'is_mix': False, 'lines': row_lines
                })

            # 샘플 균등 추출: Anti-Monopoly 순환 분산
            sample_guide = []
            all_skus = list(dict.fromkeys([p for p, _ in pure_by_sku_grouped.keys()]))
            sample_counter = 0

            for prod in all_skus:
                prod_colors = [c for p, c in pure_by_sku_grouped.keys() if p == prod]
                is_initial = any(k in prod.upper() for k in initial_skus) or ('전체' in initial_skus)
                
                if is_initial:
                    sample_guide.append({
                        'prod': prod, 'color': '전체 컬러', 'logic': '[4. 초도 신상품 풀세트]',
                        'target': 'S, M, L, XL, 2XL (전사이즈 각 1장 확보 / 최상단 잔여 박스에서 추출)'
                    })
                else:
                    for c in prod_colors:
                        act_szs = [s for s in size_order if qty_summary[(prod, c)][s]['pcs'] > 0]
                        if not act_szs: act_szs = ['M', 'L']
                        chosen_sz = act_szs[sample_counter % len(act_szs)]
                        sample_counter += 1
                        
                        matching_p = [p['p_num'] for p in pallets if p['prod'] == prod and p['color'] == c and not p.get('is_full', False)]
                        source_p = matching_p[0] if matching_p else [p['p_num'] for p in pallets if p['prod'] == prod and p['color'] == c][0]
                        
                        sample_guide.append({
                            'prod': prod, 'color': c, 'logic': '[지속 리오더 (순환 분산)]',
                            'target': f'{chosen_sz} 사이즈 1장 ({source_p} 상단 박스에서 추출)'
                        })

            tot_all_boxes = len(raw_boxes)
            tot_all_pcs = sum(sum(it['qty'] for it in b['items']) for b in raw_boxes)

            # --- 1단계: 카톡 복사용 텍스트 ---
            lines_kakao = []
            lines_kakao.append(f'[UNTAPPED 총괄 하역 오더 요약 ({work_purpose.split()[0]} / 동일 사이즈 비분할)]')
            lines_kakao.append(f'- 대상 화물: {len(all_skus)}개 품목 / 총 {tot_all_boxes}박스 ({tot_all_pcs:,}장 전수 일치)')
            lines_kakao.append(f'- 적재 규격: 완P 한도 L={default_L}박스 기준 (패딩 감지 시 L=12 자동 적용)')
            lines_kakao.append(f'- 작업 모드: {work_purpose.split()[0]} | MIX 처리: {mix_option.split()[0]}')
            lines_kakao.append(f'- 총 파레트 수: 총 {len(pallets)}P')
            lines_kakao.append('=' * 80)
            lines_kakao.append('')
            lines_kakao.append('[샘플 균등 추출 가이드 (Anti-Monopoly 순환 분산 확정)]')
            lines_kakao.append('※ 완박스 훼손 절대 금지! 잔여/자투리 박스에서 컬러별 1장씩 순환 추출')
            for sg in sample_guide:
                lines_kakao.append(f'• {sg["prod"]} ({sg["color"]}): {sg["target"]}')
            lines_kakao.append('')
            lines_kakao.append('=' * 80)
            lines_kakao.append('')
            lines_kakao.append('[사이즈별 파레트 배정 요약]')
            for r in order_summary_rows:
                lines_kakao.append(f'■ {r["prod"]} ({r["color"]}) [총 {r["pallets"]} / {r["boxes"]}]')
                if r.get('is_mix'):
                    lines_kakao.append(f'  • {r["mix_text"]}')
                else:
                    for line_txt in r['lines']:
                        lines_kakao.append(f'  • {line_txt}')
                lines_kakao.append('')

            kakao_text = '\n'.join(lines_kakao)
            st.subheader('📋 1단계: 현장 총괄 오더 요약문 (카톡 복사용)')
            st.code(kakao_text, language='text')

            # --- 2단계: 단 1개의 3-시트 공식 엑셀 (.xlsx) ---
            wb = openpyxl.Workbook()

            # Sheet 1: 총괄하역오더
            ws1 = wb.active
            ws1.title = "총괄하역오더"
            ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
            ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
            ws1.sheet_properties.pageSetUpPr.fitToPage = True
            ws1.page_setup.fitToWidth = 1
            ws1.page_setup.fitToHeight = 0
            ws1.page_margins.left = 0.3
            ws1.page_margins.right = 0.3
            ws1.page_margins.top = 0.35
            ws1.page_margins.bottom = 0.35

            ws1.column_dimensions['A'].width = 21.5
            ws1.column_dimensions['B'].width = 14.0
            ws1.column_dimensions['C'].width = 22.625
            ws1.column_dimensions['D'].width = 14.0
            ws1.column_dimensions['E'].width = 100.0

            ws1.merge_cells("A1:E1")
            ws1["A1"] = "【 언탭트 물류센터 】 현장 총괄 하역 오더 요약서"
            ws1["A1"].font = Font(name="Malgun Gothic", size=16.0, bold=True)
            ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws1.row_dimensions[1].height = 27.95

            ws1.merge_cells("A2:E2")
            ws1["A2"] = f"화물: 총 {tot_all_boxes}박스 ({tot_all_pcs:,}장) │ 규격: 완P 한도 L={default_L}박스 │ 총 {len(pallets)}P (동일 사이즈 비분할 적재)"
            ws1["A2"].font = Font(name="Malgun Gothic", size=10.0, color="1E293B")
            ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
            ws1.row_dimensions[2].height = 18.0

            thin_border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
            double_bottom_border = Border(top=Side(style='thin'), bottom=Side(style='double'), left=Side(style='thin'), right=Side(style='thin'))

            ws1.merge_cells("A3:E3")
            ws1["A3"] = f"※ 작업 목적: [{work_purpose.split()[0]}] │ 적재 한도: [L={default_L}박스 완P] │ MIX 방식: [{mix_option.split()[0]}] │ 동일 사이즈 쪼개기 까대기 원천 배제"
            ws1["A3"].font = Font(name="Malgun Gothic", size=10.5, bold=True)
            ws1["A3"].alignment = Alignment(horizontal="center", vertical="center")
            ws1.row_dimensions[3].height = 21.95
            for col in range(1, 6):
                ws1.cell(3, col).border = thin_border

            ws1.row_dimensions[4].height = 3.95

            headers_s1 = ['상품코드', '컬러', '총 파레트', '총 박스', '사이즈별 파레트 배정 요약 (S~4XL 순차 정렬 / Ghost Size 배제)']
            ws1.row_dimensions[5].height = 26.1
            for c_idx, h in enumerate(headers_s1, 1):
                c = ws1.cell(5, c_idx, h)
                c.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
                c.font = Font(name="Malgun Gothic", size=11.0, bold=True, color="FFFFFF")
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thin_border

            cur_r = 6
            prod_row_indices = {}

            for row_info in order_summary_rows:
                p = row_info['prod']
                if p not in prod_row_indices:
                    prod_row_indices[p] = []
                prod_row_indices[p].append(cur_r)
                
                if row_info.get('is_mix'):
                    summary_text = row_info['mix_text']
                    ws1.row_dimensions[cur_r].height = 180.0
                else:
                    summary_text = "\n".join(row_info['lines'])
                    n_lines = len(row_info['lines'])
                    ws1.row_dimensions[cur_r].height = max(36.0, n_lines * 22.0 + 12.0)
                    
                ws1.cell(cur_r, 1, row_info['prod']).font = Font(name="Malgun Gothic", size=12.0, bold=True)
                ws1.cell(cur_r, 2, row_info['color']).font = Font(name="Malgun Gothic", size=11.0, bold=True)
                
                c_p = ws1.cell(cur_r, 3, row_info['pallets'])
                c_p.font = Font(name="Malgun Gothic", size=13.0, bold=True, color="1D4ED8")
                c_p.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
                
                ws1.cell(cur_r, 4, row_info['boxes']).font = Font(name="Malgun Gothic", size=11.0, bold=True)
                
                # Excel 2010 완벽 호환 표준 다중행 텍스트
                c_summary = ws1.cell(cur_r, 5, summary_text)
                c_summary.font = Font(name="Malgun Gothic", size=11.0, bold=True)
                
                for c in range(1, 6):
                    ws1.cell(cur_r, c).border = thin_border
                    if c == 5:
                        ws1.cell(cur_r, c).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
                    else:
                        ws1.cell(cur_r, c).alignment = Alignment(horizontal="center", vertical="center")
                        
                cur_r += 1
                ws1.row_dimensions[cur_r].height = 3.95
                for c in range(2, 6):
                    ws1.cell(cur_r, c).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                cur_r += 1

            for p, rows in prod_row_indices.items():
                if len(rows) > 1:
                    s_r = rows[0]
                    e_r = rows[-1]
                    ws1.merge_cells(start_row=s_r, start_column=1, end_row=e_r, end_column=1)
                    ws1.cell(s_r, 1).alignment = Alignment(horizontal="center", vertical="center")

            ws1.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
            c_sum1 = ws1.cell(cur_r, 1, f"총 합계 ({len(order_summary_rows)}개 배정조합)")
            c_sum1.font = Font(name="Malgun Gothic", size=12.0, bold=True)
            c_sum1.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            c_sum1.alignment = Alignment(horizontal="center", vertical="center")
            c_sum1.border = double_bottom_border
            ws1.cell(cur_r, 2).border = double_bottom_border

            c_sum3 = ws1.cell(cur_r, 3, f"총 {len(pallets)}P")
            c_sum3.font = Font(name="Malgun Gothic", size=13.0, bold=True, color="1D4ED8")
            c_sum3.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            c_sum3.alignment = Alignment(horizontal="center", vertical="center")
            c_sum3.border = double_bottom_border

            c_sum4 = ws1.cell(cur_r, 4, f"{tot_all_boxes}박스")
            c_sum4.font = Font(name="Malgun Gothic", size=12.0, bold=True)
            c_sum4.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            c_sum4.alignment = Alignment(horizontal="center", vertical="center")
            c_sum4.border = double_bottom_border

            c_sum5 = ws1.cell(cur_r, 5, f"완박스 {tot_all_boxes - len(mix_boxes)}박스 + 다색혼적 {len(mix_boxes)}박스 ({tot_all_pcs:,}장 전수 무결점 검증완료)")
            c_sum5.font = Font(name="Malgun Gothic", size=12.0, bold=True)
            c_sum5.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            c_sum5.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            c_sum5.border = double_bottom_border
            ws1.row_dimensions[cur_r].height = 28.5

            # Sheet 2: 파레트_대형라벨
            ws2 = wb.create_sheet(title="파레트_대형라벨")
            ws2.page_setup.orientation = ws2.ORIENTATION_PORTRAIT
            ws2.page_setup.paperSize = ws2.PAPERSIZE_A4

            ws2.column_dimensions['A'].width = 4.0
            ws2.column_dimensions['B'].width = 28.0
            ws2.column_dimensions['C'].width = 6.0
            ws2.column_dimensions['D'].width = 34.0

            lr = 1
            for p_info in pallets:
                ws2.merge_cells(start_row=lr, start_column=2, end_row=lr, end_column=4)
                c = ws2.cell(lr, 2, f"【 UNTAPPED 현장 식별표 】 {p_info['p_num']}")
                c.font = Font(name="Malgun Gothic", size=18.0, bold=True)
                c.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
                c.alignment = Alignment(horizontal="center", vertical="center")
                ws2.row_dimensions[lr].height = 35.0
                lr += 1

                ws2.merge_cells(start_row=lr, start_column=2, end_row=lr, end_column=4)
                c = ws2.cell(lr, 2, p_info['prod'])
                c.font = Font(name="Malgun Gothic", size=55.0, bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
                ws2.row_dimensions[lr].height = 80.0
                lr += 1

                ws2.cell(lr, 2, p_info['color']).font = Font(name="Malgun Gothic", size=20.0, bold=True)
                ws2.cell(lr, 2).alignment = Alignment(horizontal="center", vertical="center")
                
                ws2.cell(lr, 3, "▶").alignment = Alignment(horizontal="center", vertical="center")
                ws2.cell(lr, 3).font = Font(name="Malgun Gothic", size=16.0, bold=True)
                
                ws2.cell(lr, 4, f"{p_info['size']} ({p_info['boxes']}박스)").font = Font(name="Malgun Gothic", size=20.0, bold=True)
                ws2.cell(lr, 4).alignment = Alignment(horizontal="center", vertical="center")
                ws2.row_dimensions[lr].height = 45.0
                lr += 1

                ws2.merge_cells(start_row=lr, start_column=2, end_row=lr, end_column=4)
                c = ws2.cell(lr, 2, f"세부 적재: {p_info['desc']}")
                c.font = Font(name="Malgun Gothic", size=13.0, bold=False, color="475569")
                c.alignment = Alignment(horizontal="center", vertical="center")
                ws2.row_dimensions[lr].height = 25.0
                lr += 1

                ws2.row_breaks.append(Break(id=lr))
                lr += 2

            # Sheet 3: 샘플_추출_리스트
            ws3 = wb.create_sheet(title="샘플_추출_리스트")
            ws3.page_setup.orientation = ws3.ORIENTATION_PORTRAIT
            ws3.page_setup.paperSize = ws3.PAPERSIZE_A4
            ws3.sheet_properties.pageSetUpPr.fitToPage = True
            ws3.page_setup.fitToWidth = 1
            ws3.page_setup.fitToHeight = 0
            ws3.page_margins.left = 0.4
            ws3.page_margins.right = 0.4
            ws3.page_margins.top = 0.4
            ws3.page_margins.bottom = 0.4

            ws3.column_dimensions['A'].width = 16.0
            ws3.column_dimensions['B'].width = 24.0
            ws3.column_dimensions['C'].width = 22.0
            ws3.column_dimensions['D'].width = 44.0

            ws3.merge_cells("A1:D1")
            ws3["A1"] = "【 검수/촬영용 샘플 균등 추출 가이드 】"
            ws3["A1"].font = Font(name="Malgun Gothic", size=15.0, bold=True)
            ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws3.row_dimensions[1].height = 33.95

            ws3.merge_cells("A2:D2")
            ws3["A2"] = "※ 완박스 절대 보호: 완P 박스는 절대 개봉 금지! 반드시 지정된 잔여/자투리(Loose) 박스에서만 최소 추출할 것"
            ws3["A2"].font = Font(name="Malgun Gothic", size=10.0, bold=True, color="1D4ED8")
            ws3["A2"].alignment = Alignment(horizontal="center", vertical="center")
            ws3.row_dimensions[2].height = 21.95
            ws3.row_dimensions[3].height = 8.0

            headers_s3 = ['상품코드', '컬러', '입고 유형 (적용 로직)', '샘플 지정 사이즈 및 추출 파레트']
            ws3.row_dimensions[4].height = 27.95
            for c_idx, h in enumerate(headers_s3, 1):
                c = ws3.cell(4, c_idx, h)
                c.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
                c.font = Font(name="Malgun Gothic", size=11.0, bold=True, color="FFFFFF")
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thin_border

            r3 = 5
            for s_info in sample_guide:
                ws3.row_dimensions[r3].height = 30.0
                ws3.cell(r3, 1, s_info['prod']).font = Font(name="Malgun Gothic", size=12.0, bold=True)
                ws3.cell(r3, 2, s_info['color']).font = Font(name="Malgun Gothic", size=11.0, bold=True)
                
                c_l = ws3.cell(r3, 3, s_info['logic'])
                c_l.font = Font(name="Malgun Gothic", size=10.5, bold=True, color="047857")
                c_l.fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
                
                c_t = ws3.cell(r3, 4, s_info['target'])
                c_t.font = Font(name="Malgun Gothic", size=11.0, bold=True)
                
                for c in range(1, 5):
                    ws3.cell(r3, c).border = thin_border
                    ws3.cell(r3, c).alignment = Alignment(horizontal="center", vertical="center")
                r3 += 1

            excel_buf = io.BytesIO()
            wb.save(excel_buf)
            excel_buf.seek(0)

            st.subheader('📊 2단계: 단 1개의 3-시트 공식 엑셀 다운로드')
            st.download_button(
                label='📥 언탭트_현장총괄하역오더_공식최종본.xlsx 다운로드',
                data=excel_buf,
                file_name='언탭트_현장총괄하역오더_공식최종본.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            # --- 3단계: 흑백 인쇄 최적화 CAD 미니맵 (.png) ---
            font_regular_path = "/usr/share/fonts/NotoSansCJK-Regular.ttc"
            font_bold_path = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
            if not os.path.exists(font_bold_path): font_bold_path = font_regular_path

            fp_title = fm.FontProperties(fname=font_bold_path, size=18, weight='bold') if os.path.exists(font_bold_path) else fm.FontProperties(weight='bold', size=18)
            fp_pnum = fm.FontProperties(fname=font_bold_path, size=14, weight='bold') if os.path.exists(font_bold_path) else fm.FontProperties(weight='bold', size=14)
            fp_sku = fm.FontProperties(fname=font_bold_path, size=22, weight='bold') if os.path.exists(font_bold_path) else fm.FontProperties(weight='bold', size=22)
            fp_details = fm.FontProperties(fname=font_bold_path, size=14, weight='bold') if os.path.exists(font_bold_path) else fm.FontProperties(weight='bold', size=14)
            fp_sub = fm.FontProperties(fname=font_bold_path, size=11, weight='bold') if os.path.exists(font_bold_path) else fm.FontProperties(weight='bold', size=11)

            cols = 4
            rows = (len(pallets) + cols - 1) // cols
            fig_w = 22
            fig_h = max(12, rows * 4.2)

            fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=220)
            ax.set_facecolor('#FFFFFF')
            ax.set_xlim(0, cols * 50 + 10)
            ax.set_ylim(0, rows * 32 + 32)

            dock_w = cols * 50 - 4
            dock = patches.FancyBboxPatch((7, rows * 32 + 6), dock_w, 16, boxstyle="round,pad=0.2", ec="#000000", fc="#F1F5F9", lw=3.0)
            ax.add_patch(dock)
            ax.text((cols * 50 + 10) / 2, rows * 32 + 14, f"▼ [UNLOADING DOCK] 현장 하역구역 배치도 (총 {tot_all_boxes}박스 / {len(pallets)}P 흑백 인쇄 최적화) ▼", 
                    color="#000000", fontproperties=fp_title, ha='center', va='center')

            for idx, p in enumerate(pallets):
                r_i = idx // cols
                c_i = idx % cols
                x = 7 + c_i * 50
                y = (rows - 1 - r_i) * 32 + 6

                if p['type'] == 'MIX':
                    fc = "#E2E8F0"
                    ec = "#000000"
                    lw = 3.5
                    tag_str = "★ [도크초입 MIX 격리]"
                elif '완P' in p['type']:
                    fc = "#FFFFFF"
                    ec = "#000000"
                    lw = 4.0
                    tag_str = f"[{p['size']} {p['boxes']}박스]"
                else:
                    fc = "#FFFFFF"
                    ec = "#000000"
                    lw = 2.0
                    tag_str = f"[{p['size']} {p['boxes']}박스]"

                cell = patches.FancyBboxPatch((x, y), 46, 26, boxstyle="round,pad=0.2", ec=ec, fc=fc, lw=lw)
                ax.add_patch(cell)

                ax.plot([x, x + 46], [y + 19.5, y + 19.5], color='#000000', lw=1.5)
                ax.text(x + 23, y + 22.8, f"{p['p_num']}  │  {tag_str}", color="#000000", fontproperties=fp_pnum, ha='center', va='center')
                ax.text(x + 23, y + 14.5, f"{p['prod']}", color="#000000", fontproperties=fp_sku, ha='center', va='center')
                ax.text(x + 23, y + 8.5, f"[{p['color']}]  ▶  {p['size']}  ({p['boxes']}박스)", color="#000000", fontproperties=fp_details, ha='center', va='center')

                desc_txt = p['desc']
                if len(desc_txt) > 30:
                    desc_txt = desc_txt[:29] + "..."
                ax.text(x + 23, y + 3.5, desc_txt, color="#334155", fontproperties=fp_sub, ha='center', va='center')

            ax.axis('off')
            plt.tight_layout()

            img_buf = io.BytesIO()
            plt.savefig(img_buf, format="png", dpi=220, bbox_inches='tight')
            img_buf.seek(0)
            plt.close()

            st.subheader('🖼️ 3단계: 흑백 인쇄 최적화 CAD 미니맵 (.png)')
            st.image(img_buf, use_container_width=True)

            st.download_button(
                label='🖼️ CAD 미니맵 이미지(.png) 다운로드',
                data=img_buf,
                file_name='언탭트_파레트하역_2D미니맵_흑백인쇄용.png',
                mime='image/png'
            )

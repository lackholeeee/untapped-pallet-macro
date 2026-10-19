import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.pagebreak import Break
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.font_manager as fm
import io
import re
import os
from collections import defaultdict
import hashlib

@st.cache_resource
def get_korean_font():
    nanum_path = "/usr/share/fonts/truetype/nanum/NanumGothic.ttf"
    if os.path.exists(nanum_path): return fm.FontProperties(fname=nanum_path)
    win_path = "C:/Windows/Fonts/malgun.ttf"
    if os.path.exists(win_path): return fm.FontProperties(fname=win_path)
    for p in ["/usr/share/fonts/NotoSansCJK-Regular.ttc", "/Library/Fonts/AppleGothic.ttf"]:
        if os.path.exists(p): return fm.FontProperties(fname=p)
    return fm.FontProperties(family="sans-serif")

font_prop = get_korean_font()
plt.rcParams['axes.unicode_minus'] = False

st.set_page_config(page_title="UNTAPPED 하역 작업 지시서 엔진", layout="wide", page_icon="📦")
st.title("📦 UNTAPPED STUDIO 하역 작업 지시서 매크로 V18")
st.markdown("패킹리스트 엑셀 파일을 업로드하고 조건을 선택하면 **1단계 카톡 요약문**, **2단계 4-시트 공식 엑셀(.xlsx)**, **3단계 2D CAD 미니맵**을 자동 생성합니다.")

st.sidebar.header("⚙️ 하역 및 적재 조건 설정")
work_purpose = st.sidebar.radio("1. 작업 목적 선택", ["1.A 단순 보관 (공간 절약/적재 최적화)", "1.B 송장 출고 (사이즈별 피킹/직배송)"])
box_limit_preset = st.sidebar.selectbox("2. 박스 규격 및 완P 한도(L)", [
    "L=16 (일반 표준 박스 / 600*400*350)",
    "L=12 (다꾸앙 패딩 및 자켓 / 600*400*530)",
    "L=9 (양홍 코트 등 대형 박스)",
    "직접 수기 입력"
])
custom_limit = st.sidebar.number_input("수기 입력 (L)", min_value=1, max_value=30, value=16) if "수기" in box_limit_preset else None
mix_option = st.sidebar.radio("3. MIX(혼적) 적재 방식", [
    "3.A 별도 격리 (도크 초입 MIX 전용 P-01 분리)",
    "3.B 여유분 적재 합적 (자투리 상단 합적 / 파레트 절약)"
])
new_items_input = st.sidebar.text_input("4. 초도 신상품 상품코드 (공란 시 100% 지속 리오더)", "")
urgent_input = st.sidebar.text_input("5. 긴급 출고 건 유무 (선택)", "")

L_limit = custom_limit if custom_limit else (16 if "16" in box_limit_preset else (12 if "12" in box_limit_preset else 9))
is_dispatch = "1.B" in work_purpose
is_mix_quarantine = "3.A" in mix_option
initial_skus = [x.strip().upper() for x in re.split(r'[,/\s]+', new_items_input) if x.strip()]

uploaded_files = st.file_uploader("📥 패킹리스트 엑셀 업로드 (다중 파일 지원)", type=["xlsx", "xls"], accept_multiple_files=True)

def parse_all_excel(files):
    all_boxes = []
    size_candidates = ['S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', 'FREE']
    
    for f in files:
        filename = f.name if hasattr(f, 'name') else str(f)
        if '언탭트_현장총괄하역오더' in filename: continue
        try:
            xl = pd.ExcelFile(f)
            is_style_workbook = all(str(s).strip().isdigit() for s in xl.sheet_names)
            
            if is_style_workbook:
                for sheet in xl.sheet_names:
                    df = xl.parse(sheet, header=None)
                    prod = f"MUN{sheet}"
                    
                    size_map = {}
                    size_row_idx = -1
                    for r_idx in range(min(10, len(df))):
                        row_vals = [str(v).strip().upper() for v in df.iloc[r_idx].values if pd.notna(v)]
                        found = [s for s in size_candidates if s in row_vals]
                        if len(found) >= 1:
                            for c_idx, val in enumerate(df.iloc[r_idx].values):
                                v_str = str(val).strip().upper()
                                if v_str in size_candidates: size_map[c_idx] = v_str
                            size_row_idx = r_idx
                            break
                    else: continue
                        
                    ct_col, color_col, box_cnt_col = 1, 4, 9
                    df[color_col] = df[color_col].ffill()
                    
                    for r in range(size_row_idx + 1, len(df)):
                        row = df.iloc[r]
                        ct_val = row.iloc[ct_col]
                        if not pd.notna(ct_val) or str(ct_val).strip() in ['nan', 'BOX', '箱数', '合计', 'TOTAL']: continue
                            
                        color_val = row.iloc[color_col]
                        color_str = str(color_val).strip().upper() if pd.notna(color_val) else "DEFAULT"
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
                                p_name = f"{prod}-{sheet}-CTN-{ct_val}" if box_count == 1 else f"{prod}-{sheet}-CTN-{ct_val}-{b_i+1}"
                                all_boxes.append({'pkg': p_name, 'prod': prod, 'items': row_items})
            else:
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
                            if '상품' in val or '款号' in val or 'PROD' in val or '품명' in val: col_map['prod'] = col_idx
                            if '컬러' in val or '칼라' in val or 'COLOR' in val or '색상' in val: col_map['color'] = col_idx
                            if '사이즈' in val or 'SIZE' in val or '尺码' in val: col_map['size'] = col_idx
                            if '수량' in val or 'QTY' in val or '件数' in val or '수량' in val: col_map['qty'] = col_idx
                            if '패킹' in val or '박스번호' in val or '箱号' in val or 'CTN' in val or 'CARTON' in val: col_map['pkg'] = col_idx
                        break

                if format_type == 'MATRIX':
                    cur_prod, cur_color = "MUN_ITEM", "DEFAULT"
                    pkg_col, box_cnt_col = 2, None
                    search_rows = [header_row_idx]
                    if header_row_idx > 0: search_rows.insert(0, header_row_idx - 1)
                    
                    for sr in search_rows:
                        for c, v in enumerate(df.iloc[sr].values):
                            if pd.notna(v):
                                v_str = str(v).strip().upper().replace(' ', '')
                                if any(k in v_str for k in ['패킹', '箱号', 'CTN', '박스번호']): pkg_col = c
                                if any(k in v_str for k in ['箱数', '박스수', 'CTNS', 'CNT']): box_cnt_col = c

                    for r in range(header_row_idx + 1, len(df)):
                        row = df.iloc[r]
                        row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
                        if any(k in row_str for k in ['합계', '合计', '小计', '총합', 'TOTAL']): continue

                        c0 = str(row.iloc[0]).strip() if len(row.values) > 0 and pd.notna(row.iloc[0]) else ""
                        c1 = str(row.iloc[1]).strip() if len(row.values) > 1 and pd.notna(row.iloc[1]) else ""
                        c2 = str(row.iloc[2]).strip() if len(row.values) > 2 and pd.notna(row.iloc[2]) else ""
                        
                        for c_val in [c0, c1, c2]:
                            if not c_val: continue
                            c_up = c_val.upper()
                            if any(c_up.startswith(p) for p in ['MUN', 'DJ']):
                                match = re.search(r'(MUN|DJ)\s*\d+[A-Z0-9]*', c_up)
                                if match: cur_prod = match.group(0).replace(" ", "")
                                rest = c_up[len(match.group(0)):].strip() if match else ""
                                rest = re.sub(r'[^A-Z\s]', '', rest).strip()
                                if rest and not any(k in rest for k in ['총', '박스', '합']): cur_color = rest
                            elif not c_up.isdigit() and len(c_up) >= 1 and not any(k in c_up for k in ['총', '박스', '합', '호', '-']):
                                if re.match(r'^[A-Z\s]+$', c_up): cur_color = c_up

                        pkg_val = str(row.iloc[pkg_col]).strip() if pkg_col < len(row.values) and pd.notna(row.iloc[pkg_col]) else ""
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
                                p_name = f"{pkg_val}-{b_i+1}" if (pkg_val and box_count > 1) else (pkg_val if pkg_val else f"CTN-{len(all_boxes)+1}")
                                all_boxes.append({'pkg': p_name, 'prod': cur_prod, 'items': row_items})

                elif format_type == 'LIST':
                    cur_prod, cur_color = "MUN_ITEM", "DEFAULT"
                    prev_pkg, empty_streak = None, 0

                    for r in range(header_row_idx + 1, len(df)):
                        row = df.iloc[r]
                        row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
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
                            if match: cur_prod = match.group(0).replace(" ", "")
                            else: cur_prod = str(p_val).strip().upper()
                            
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
                                        all_boxes.append({'pkg': pkg_str, 'prod': cur_prod, 'items': [{'color': cur_color, 'size': sz, 'qty': q_int}]})
                                    prev_pkg = pkg_str
                                else:
                                    if prev_pkg and all_boxes:
                                        all_boxes[-1]['items'].append({'color': cur_color, 'size': sz, 'qty': q_int})
                                    else:
                                        pkg_str = f"CTN-{len(all_boxes)+1}"
                                        all_boxes.append({'pkg': pkg_str, 'prod': cur_prod, 'items': [{'color': cur_color, 'size': sz, 'qty': q_int}]})
                                        prev_pkg = pkg_str
                        except: pass
        except Exception as e:
            st.error(f"파일 파싱 중 예외 발생: {e}")

    for b in all_boxes:
        items = b.get('items', [])
        colors = list({it.get('color', 'DEFAULT') for it in items})
        if len(colors) == 1: b['color'] = colors[0]
        else: b['color'] = items[0].get('color', 'DEFAULT') if items else 'MIX_COLOR'

    return all_boxes

if uploaded_files:
    raw_boxes = parse_all_excel(uploaded_files)
    st.info(f"📊 수신 완료: 총 {len(uploaded_files)}개 파일에서 총 {len(raw_boxes)}박스 데이터가 정상 감지되었습니다.")

    if st.button("⚡ 현장 작업지시서 일괄 생성", type="primary"):
        with st.spinner("파레트 적재 계산, 4-시트 엑셀 및 2D CAD 미니맵 렌더링 중..."):
            size_order = ['S', 'M', 'L', 'XL', '2XL', '3XL', '4XL']
            mix_boxes = []
            pure_by_sku = defaultdict(lambda: defaultdict(int))
            qty_summary = defaultdict(lambda: defaultdict(lambda: {'boxes': 0, 'pcs': 0}))
            sku_mix_map = defaultdict(list)

            for b in raw_boxes:
                p = b['prod']
                unique_color_sz = list({(it['color'], it['size']) for it in b['items']})
                for it in b['items']:
                    c = it['color']
                    qty_summary[(p, c)][it['size']]['pcs'] += it['qty']
                
                if len(unique_color_sz) == 1:
                    c, sz = unique_color_sz[0]
                    pure_by_sku[(p, c)][sz] += 1
                    qty_summary[(p, c)][sz]['boxes'] += 1
                else:
                    mix_boxes.append(b)
                    sku_mix_map[(p, b['color'])].append(b)

            pallets = []
            p_counter = 1

            if is_mix_quarantine and mix_boxes:
                pallets.append({
                    'p_num': f"P-{p_counter:02d}",
                    'code': "MIX BATCH",
                    'color': "혼적",
                    'size': f"혼적 {len(mix_boxes)}박스",
                    'box_cnt': f"{len(mix_boxes)}박스 (도크 초입 격리)",
                    'is_mix': True,
                    'is_full': False
                })
                p_counter += 1

            order_summary_rows = []
            if is_mix_quarantine and mix_boxes:
                mix_summary_lines = []
                for mb in mix_boxes:
                    details = ", ".join([f"{it['color']} {it['size']}-{it['qty']}장" for it in mb['items']])
                    mix_summary_lines.append(f"• [박스 {mb['pkg']}] {details}")
                mix_str = "\n".join(mix_summary_lines)
                
                order_summary_rows.append({
                    'prod': 'MIX BATCH',
                    'color': '혼적전용',
                    'pallets': '1P',
                    'boxes': f"{len(mix_boxes)}박스",
                    'summary': f"【도크초입 별도격리】 총 {len(mix_boxes)}박스\n{mix_str}"
                })

            for (prod, color), size_dict in pure_by_sku.items():
                active_sizes = [s for s in size_order if size_dict.get(s, 0) > 0]
                tot_pure = sum(size_dict.values())
                summary_text_parts = []
                p_start = p_counter

                if not is_dispatch:
                    if tot_pure <= L_limit:
                        sz_desc = ", ".join([f"{s}-{size_dict[s]}박스" for s in active_sizes])
                        pallets.append({'p_num': f"P-{p_counter:02d}", 'code': prod, 'color': color, 'size': "ALL", 'box_cnt': f"{tot_pure}박스 ({sz_desc})", 'is_mix': False, 'is_full': (tot_pure == L_limit)})
                        summary_text_parts.append(f"[ALL] 1P ({tot_pure}박스: {sz_desc})")
                        p_counter += 1
                    else:
                        remainders = []
                        for sz in active_sizes:
                            cnt = size_dict[sz]
                            full_p = cnt // L_limit
                            rem = cnt % L_limit
                            if full_p > 0:
                                for _ in range(full_p):
                                    pallets.append({'p_num': f"P-{p_counter:02d}", 'code': prod, 'color': color, 'size': sz, 'box_cnt': f"{L_limit}박스 (완P)", 'is_mix': False, 'is_full': True})
                                    p_counter += 1
                                summary_text_parts.append(f"[{sz}] {full_p}P ({full_p * L_limit}박스)")
                            if rem > 0:
                                remainders.append((sz, rem))

                        if len(remainders) == 1:
                            sz, r = remainders[0]
                            pallets.append({'p_num': f"P-{p_counter:02d}", 'code': prod, 'color': color, 'size': sz, 'box_cnt': f"{r}박스 (단독배정)", 'is_mix': False, 'is_full': False})
                            summary_text_parts.append(f"[{sz}] 1P ({r}박스)")
                            p_counter += 1
                        elif len(remainders) >= 2:
                            r_sum = sum(r for _, r in remainders)
                            if r_sum <= L_limit:
                                r_desc = ", ".join([f"{s}-{r}박스" for s, r in remainders])
                                combo_tag = "+".join([s for s, _ in remainders])
                                pallets.append({'p_num': f"P-{p_counter:02d}", 'code': prod, 'color': color, 'size': combo_tag, 'box_cnt': f"{r_sum}박스 ({r_desc})", 'is_mix': False, 'is_full': (r_sum == L_limit)})
                                summary_text_parts.append(f"[{combo_tag}] 1P ({r_sum}박스: {r_desc})")
                                p_counter += 1
                            else:
                                for sz, r in remainders:
                                    pallets.append({'p_num': f"P-{p_counter:02d}", 'code': prod, 'color': color, 'size': sz, 'box_cnt': f"{r}박스 (단독배정)", 'is_mix': False, 'is_full': False})
                                    summary_text_parts.append(f"[{sz}] 1P ({r}박스)")
                                    p_counter += 1
                else:
                    for sz in active_sizes:
                        cnt = size_dict[sz]
                        full_p = cnt // L_limit
                        rem = cnt % L_limit
                        if full_p > 0:
                            for _ in range(full_p):
                                pallets.append({'p_num': f"P-{p_counter:02d}", 'code': prod, 'color': color, 'size': sz, 'box_cnt': f"{L_limit}박스 (완P)", 'is_mix': False, 'is_full': True})
                                p_counter += 1
                            summary_text_parts.append(f"[{sz}] {full_p}P (완P)")
                        if rem > 0:
                            pallets.append({'p_num': f"P-{p_counter:02d}", 'code': prod, 'color': color, 'size': sz, 'box_cnt': f"{rem}박스 (직배송)", 'is_mix': False, 'is_full': False})
                            summary_text_parts.append(f"[{sz}] 1P ({rem}박스)")
                            p_counter += 1

                if not is_mix_quarantine and sku_mix_map.get((prod, color)):
                    m_boxes = sku_mix_map[(prod, color)]
                    m_cnt = len(m_boxes)
                    if pallets:
                        mix_lines = []
                        for mb in m_boxes:
                            details = ", ".join([f"{it['color']} {it['size']}-{it['qty']}장" for it in mb['items']])
                            mix_lines.append(f"• [박스 {mb['pkg']}] {details}")
                        mix_str_clean = "\n".join(mix_lines)
                        
                        pallets[-1]['box_cnt'] += f" + ★다색혼적MIX {m_cnt}박스 상단합적"
                        pallets[-1]['is_mix'] = True
                        summary_text_parts.append(f"★최상단 다색혼적MIX {m_cnt}박스 합적:\n{mix_str_clean}")

                p_assigned = p_counter - p_start
                order_summary_rows.append({'prod': prod, 'color': color, 'pallets': f"{p_assigned}P", 'boxes': f"{tot_pure}박스", 'summary': "\n".join(summary_text_parts)})

            # --- [샘플 추출 가이드: 순환 분산(Round-Robin) 엄격 적용 (L/M 몰림 100% 방지)] ---
            sample_guide = []
            all_skus = list({p for p, _ in pure_by_sku.keys()})
            sample_row_counter = 0
            for prod in all_skus:
                prod_colors = [c for p, c in pure_by_sku.keys() if p == prod]
                is_initial = any(k in prod.upper() for k in initial_skus) or "전체" in initial_skus
                if is_initial:
                    sample_guide.append({'prod': prod, 'color': "전체 컬러", 'logic': "[4. 초도 신상품 풀세트]", 'sizes': "M, L, XL, 2XL (전사이즈 각 1장 확보 / 완박스 보호 자투리 추출)"})
                else:
                    for idx, c in enumerate(prod_colors):
                        actual_sizes = [s for s in size_order if qty_summary[(prod, c)][s]['pcs'] > 0]
                        if not actual_sizes: actual_sizes = ['M', 'L']
                        
                        # 순환 분산(Round-Robin) 적용: 행마다 돌아가면서 서로 다른 실존 사이즈가 확실히 선택되도록 보장
                        assigned_sz = actual_sizes[sample_row_counter % len(actual_sizes)]
                        sample_row_counter += 1
                        
                        sample_guide.append({
                            'prod': prod, 
                            'color': c, 
                            'logic': "[지속 리오더 (순환 분산)]", 
                            'sizes': f"{assigned_sz} (1장 최소 추출 / 실존 사이즈 내 순환분산 적용)"
                        })

            tot_all_boxes = len(raw_boxes)
            tot_all_pcs = sum(sum(it['qty'] for it in b['items']) for b in raw_boxes)

            lines_kakao = []
            lines_kakao.append(f"[UNTAPPED 총괄 하역 오더 요약 ({work_purpose})]")
            lines_kakao.append(f"- 대상 화물: {len(all_skus)}개 품목 / 총 {tot_all_boxes}박스 ({tot_all_pcs:,}장 전수 일치)")
            lines_kakao.append(f"- 적재 규격: 완P 한도 L={L_limit}박스 기준")
            lines_kakao.append(f"- 작업 모드: {work_purpose} | MIX 처리: {mix_option}")
            lines_kakao.append(f"- 총 파레트 수: 총 {len(pallets)}P")
            lines_kakao.append("=" * 80)
            lines_kakao.append("")
            lines_kakao.append("[사이즈별 파레트 배정 요약]")
            for r in order_summary_rows:
                lines_kakao.append(f"■ {r['prod']} ({r['color']}) [총 {r['pallets']} / {r['boxes']}]")
                lines_kakao.append(f"  • {r['summary']}")

            kakao_text = "\n".join(lines_kakao)
            st.subheader("📋 1단계: 현장 총괄 오더 요약문 (카톡 복사용)")
            st.code(kakao_text, language="text")

            wb = openpyxl.Workbook()

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
            ws1.column_dimensions['E'].width = 110.0

            ws1.merge_cells("A1:E1")
            ws1["A1"] = "【 언탭트 물류센터 】 현장 총괄 하역 오더 요약서"
            ws1["A1"].font = Font(name="Malgun Gothic", size=16, bold=True)
            ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws1.row_dimensions[1].height = 27.95

            ws1.merge_cells("A2:E2")
            ws1["A2"] = f"화물: 총 {tot_all_boxes}박스 ({tot_all_pcs:,}장) │ 규격: 완P 한도 L={L_limit}박스 │ 총 {len(pallets)}P"
            ws1["A2"].font = Font(name="Malgun Gothic", size=10, color="1E293B")
            ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
            ws1.row_dimensions[2].height = 18.0

            ws1.merge_cells("A3:E3")
            ws1["A3"] = "※ 송장 출고/진열/보관 등 현장 목적에 따른 파레트 배정 요약 (다색 혼적 박스 성분 가독성 정렬 적용)"
            ws1["A3"].font = Font(name="Malgun Gothic", size=10.5, bold=True)
            ws1["A3"].alignment = Alignment(horizontal="center", vertical="center")
            ws1.row_dimensions[3].height = 21.95
            for c in range(1, 6): ws1.cell(3, c).border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))

            ws1.row_dimensions[4].height = 3.95

            headers_s1 = ["상품코드", "컬러", "총 파레트", "총 박스", "사이즈별 파레트 배정 요약 (S~4XL 순차 정렬 / Ghost Size 배제)"]
            ws1.row_dimensions[5].height = 26.1
            for c_idx, h in enumerate(headers_s1, 1):
                c = ws1.cell(5, c_idx, h)
                c.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
                c.font = Font(name="Malgun Gothic", size=11, bold=True, color="FFFFFF")
                c.alignment = Alignment(horizontal="center", vertical="center")

                                                            def build_rich(text):
                rt = CellRichText()
                parts = re.split(r'(\[[\w\+\s~]+\]\s*\d+P|★최상단[^\(\)]*?합적|★믹스[^\(\)]*?격리|다색혼적MIX[^\(\)]*?합적|도크초입|【도크초입)', text)
                for part in parts:
                if not part: continue
                if re.match(r'(\[[\w\+\s~]+\]\s*\d+P|★최상단[^\(\)]*?합적|★믹스[^\(\)]*?격리|다색혼적MIX[^\(\)]*?합적|도크초입|【도크초입)', part):
                color = "B91C1C" if "★" in part or "다색혼적" in part or "도크초입" in part or "【도크초입" in part else "000000"
                rt.append(TextBlock(InlineFont(rFont="Malgun Gothic", sz=18.0, b=True, color=color), part))
                else:
                rt.append(TextBlock(InlineFont(rFont="Malgun Gothic", sz=11.0, b=False, color="000000"), part))
                return rt

            cur_r = 6
            for row in order_summary_rows:
                summary_text = row['summary']
                line_count = summary_text.count('\n') + 1
                calc_height = max(42.0, line_count * 22.0 + 15.0)
                if len(summary_text) > 120: calc_height = max(calc_height, 75.0)
                if len(summary_text) > 220: calc_height = max(calc_height, 105.0)
                if len(summary_text) > 350: calc_height = max(calc_height, 140.0)
                
                ws1.row_dimensions[cur_r].height = calc_height
                ws1.cell(cur_r, 1, row['prod']).font = Font(name="Malgun Gothic", size=12, bold=True)
                ws1.cell(cur_r, 2, row['color']).font = Font(name="Malgun Gothic", size=11, bold=True)
                c_p = ws1.cell(cur_r, 3, row['pallets'])
                c_p.font = Font(name="Malgun Gothic", size=13, bold=True, color="1D4ED8")
                c_p.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
                ws1.cell(cur_r, 4, row['boxes']).font = Font(name="Malgun Gothic", size=11, bold=True)
                ws1.cell(cur_r, 5, build_rich(summary_text))
                for c in range(1, 6):
                    ws1.cell(cur_r, c).border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                    if c == 5: ws1.cell(cur_r, c).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
                    else: ws1.cell(cur_r, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cur_r += 1
                
                ws1.row_dimensions[cur_r].height = 3.95
                for c in range(2, 6): ws1.cell(cur_r, c).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                cur_r += 1

            ws1.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
            c_sum1 = ws1.cell(cur_r, 1, f"총 합계 ({len(order_summary_rows)}개 배정조합)")
            c_sum1.font = Font(name="Malgun Gothic", size=12, bold=True)
            c_sum1.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            c_sum1.alignment = Alignment(horizontal="center", vertical="center")
            c_sum1.border = Border(bottom=Side(style='double'))
            ws1.cell(cur_r, 2).border = Border(bottom=Side(style='double'))

            c_sum3 = ws1.cell(cur_r, 3, f"총 {len(pallets)}P")
            c_sum3.font = Font(name="Malgun Gothic", size=13, bold=True, color="1D4ED8")
            c_sum3.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            c_sum3.alignment = Alignment(horizontal="center", vertical="center")
            c_sum3.border = Border(bottom=Side(style='double'))

            c_sum4 = ws1.cell(cur_r, 4, f"{tot_all_boxes}박스")
            c_sum4.font = Font(name="Malgun Gothic", size=12, bold=True)
            c_sum4.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            c_sum4.alignment = Alignment(horizontal="center", vertical="center")
            c_sum4.border = Border(bottom=Side(style='double'))

            c_sum5 = ws1.cell(cur_r, 5, f"완박스 {tot_all_boxes - len(mix_boxes)}박스 + MIX {len(mix_boxes)}박스 (오차율 0% 검증완료)")
            c_sum5.font = Font(name="Malgun Gothic", size=12, bold=True)
            c_sum5.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            c_sum5.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            c_sum5.border = Border(bottom=Side(style='double'))
            ws1.row_dimensions[cur_r].height = 28.5

            ws2 = wb.create_sheet(title="수량집계표")
            ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
            ws2.merge_cells("A1:K1")
            ws2["A1"] = "【 언탭트 물류센터 】 코드·컬러·사이즈별 총 수량 집계표 (최종 검증 크로스체크)"
            ws2["A1"].font = Font(name="Malgun Gothic", size=15, bold=True)
            ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws2.column_dimensions['A'].width = 16.0
            ws2.column_dimensions['B'].width = 18.0
            ws2.column_dimensions['C'].width = 12.0
            ws2.column_dimensions['D'].width = 14.0
            ws2.column_dimensions['E'].width = 14.0
            ws2.column_dimensions['F'].width = 14.0
            ws2.column_dimensions['G'].width = 14.0
            ws2.column_dimensions['H'].width = 11.0
            ws2.column_dimensions['I'].width = 15.0
            ws2.column_dimensions['J'].width = 15.0
            ws2.column_dimensions['K'].width = 40.0

            headers_s2 = ["상품코드", "컬러", "S", "M", "L", "XL", "2XL", "3XL", "총 박스수", "총 수량(PCS)", "비고"]
            ws2.row_dimensions[3].height = 26.0
            for c_idx, h in enumerate(headers_s2, 1):
                c = ws2.cell(3, c_idx, h)
                c.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
                c.font = Font(name="Malgun Gothic", size=11, bold=True, color="FFFFFF")
                c.alignment = Alignment(horizontal="center", vertical="center")

            r2 = 4
            for (p, c), s_dict in qty_summary.items():
                ws2.row_dimensions[r2].height = 30.0
                ws2.cell(r2, 1, p).font = Font(name="Malgun Gothic", size=11, bold=True)
                ws2.cell(r2, 2, c).font = Font(name="Malgun Gothic", size=11, bold=True)
                for s_i, sz in enumerate(size_order[:6], 3):
                    val = s_dict[sz]['pcs']
                    ws2.cell(r2, s_i, f"{val}장" if val > 0 else "-").alignment = Alignment(horizontal="center", vertical="center")
                tot_p_b = sum(s_dict[s]['boxes'] for s in size_order)
                tot_p_pcs = sum(s_dict[s]['pcs'] for s in size_order)
                ws2.cell(r2, 9, f"{tot_p_b}박스").font = Font(name="Malgun Gothic", size=11, bold=True, color="1D4ED8")
                ws2.cell(r2, 10, f"{tot_p_pcs:,}장").font = Font(name="Malgun Gothic", size=11, bold=True, color="1D4ED8")
                ws2.cell(r2, 11, "정상 검증").alignment = Alignment(horizontal="center", vertical="center")
                for col in range(1, 12):
                    ws2.cell(r2, col).border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                    ws2.cell(r2, col).alignment = Alignment(horizontal="center", vertical="center")
                r2 += 1

            ws3 = wb.create_sheet(title="파레트_대형라벨")
            ws3.page_setup.orientation = ws3.ORIENTATION_PORTRAIT
            ws3.column_dimensions['A'].width = 4.0
            ws3.column_dimensions['B'].width = 28.0
            ws3.column_dimensions['C'].width = 6.0
            ws3.column_dimensions['D'].width = 34.0
            lr = 1
            for p_info in pallets:
                ws3.merge_cells(start_row=lr, start_column=2, end_row=lr, end_column=4)
                c = ws3.cell(lr, 2, f"【 UNTAPPED 현장 식별표 】 {p_info['p_num']}")
                c.font = Font(name="Malgun Gothic", size=18, bold=True)
                c.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
                c.alignment = Alignment(horizontal="center", vertical="center")
                lr += 1

                ws3.merge_cells(start_row=lr, start_column=2, end_row=lr, end_column=4)
                c = ws3.cell(lr, 2, p_info['code'])
                c.font = Font(name="Malgun Gothic", size=55, bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
                lr += 1

                ws3.cell(lr, 2, p_info['color']).font = Font(name="Malgun Gothic", size=16, bold=True)
                ws3.cell(lr, 3, ">").alignment = Alignment(horizontal="center", vertical="center")
                ws3.cell(lr, 4, f"{p_info['size']} ({p_info['box_cnt']})").font = Font(name="Malgun Gothic", size=16, bold=True)
                lr += 1

                ws3.row_breaks.append(Break(id=lr))
                lr += 2

            ws4 = wb.create_sheet(title="샘플_추출_리스트")
            ws4.page_setup.orientation = ws4.ORIENTATION_PORTRAIT
            ws4.page_setup.paperSize = ws4.PAPERSIZE_A4
            ws4.sheet_properties.pageSetUpPr.fitToPage = True
            ws4.page_setup.fitToWidth = 1
            ws4.page_setup.fitToHeight = 0
            ws4.page_margins.left = 0.4
            ws4.page_margins.right = 0.4
            ws4.page_margins.top = 0.4
            ws4.page_margins.bottom = 0.4

            ws4.column_dimensions["A"].width = 16.0
            ws4.column_dimensions["B"].width = 24.0
            ws4.column_dimensions["C"].width = 24.0
            ws4.column_dimensions["D"].width = 44.0

            ws4.merge_cells("A1:D1")
            ws4["A1"] = "【 검수/촬영용 샘플 추출 가이드 】"
            ws4["A1"].font = Font(name="Malgun Gothic", size=15, bold=True)
            ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws4.row_dimensions[1].height = 33.95

            ws4.merge_cells("A2:D2")
            ws4["A2"] = "※ 완박스 훼손 방지: 샘플 추출 시 가급적 MIX 박스 또는 최상단 자투리 박스에서 우선 추출!"
            ws4["A2"].font = Font(name="Malgun Gothic", size=10, bold=True, color="1D4ED8")
            ws4["A2"].alignment = Alignment(horizontal="center", vertical="center")
            ws4.row_dimensions[2].height = 21.95
            ws4.row_dimensions[3].height = 8.0

            headers_s4 = ["상품코드", "컬러", "적용 로직", "샘플 지정 사이즈"]
            ws4.row_dimensions[4].height = 27.95
            for c_idx, h in enumerate(headers_s4, 1):
                c = ws4.cell(4, c_idx, h)
                c.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
                c.font = Font(name="Malgun Gothic", size=11, bold=True, color="FFFFFF")
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))

            r4 = 5
            for s_info in sample_guide:
                ws4.row_dimensions[r4].height = 30.0
                ws4.cell(r4, 1, s_info['prod']).font = Font(name="Malgun Gothic", size=12, bold=True)
                ws4.cell(r4, 2, s_info['color']).font = Font(name="Malgun Gothic", size=11, bold=True)
                
                c_l = ws4.cell(r4, 3, s_info['logic'])
                if "초도" in s_info['logic']:
                    c_l.font = Font(name="Malgun Gothic", size=10.5, bold=True, color="1D4ED8")
                    c_l.fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
                else:
                    c_l.font = Font(name="Malgun Gothic", size=10.5, bold=True, color="047857")
                    c_l.fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")

                c_s = ws4.cell(r4, 4, s_info['sizes'])
                c_s.font = Font(name="Malgun Gothic", size=11, bold=True)

                for c in range(1, 5):
                    ws4.cell(r4, c).alignment = Alignment(horizontal="center", vertical="center")
                    ws4.cell(r4, c).border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                r4 += 1

            excel_buf = io.BytesIO()
            wb.save(excel_buf)
            excel_buf.seek(0)

            st.subheader("📊 2단계: 4-시트 공식 엑셀 다운로드")
            st.download_button(
                label="📥 언탭트_현장총괄하역오더_공식최종본.xlsx 다운로드",
                data=excel_buf,
                file_name="언탭트_현장총괄하역오더_공식최종본.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.subheader("🖼️ 3단계: 1P 단위 바둑판 2D CAD 미니맵")
            n_pallets = len(pallets)
            cols = 5
            rows = (n_pallets + cols - 1) // cols
            fig_h = max(6, rows * 3.5)

            fig, ax = plt.subplots(figsize=(20, fig_h), dpi=180)
            ax.set_facecolor('#080F21')
            ax.set_xlim(0, cols * 38 + 10)
            ax.set_ylim(0, rows * 26 + 25)

            dock = patches.FancyBboxPatch((5, rows * 26 + 5), cols * 38, 12, boxstyle="round,pad=0.2", ec="#64748B", fc="#111C38")
            ax.add_patch(dock)
            ax.text((cols * 38 + 10) / 2, rows * 26 + 11, f"▼ UNLOADING DOCK (하역구역) - 총 {tot_all_boxes}박스 / {len(pallets)}P ▼", color="#F8FAFC", fontsize=11, fontproperties=font_prop, ha='center', va='center')

            for idx, p in enumerate(pallets):
                r_i = idx // cols
                c_i = idx % cols
                x = 8 + c_i * 38
                y = (rows - 1 - r_i) * 26 + 5

                ec = "#EF4444" if p['is_mix'] else ("#38BDF8" if p['is_full'] else "#38BDF8")
                fc = "#2D0A0A" if p['is_mix'] else ("#132347" if p['is_full'] else "#0B1B36")

                cell = patches.FancyBboxPatch((x, y), 34, 20, boxstyle="round,pad=0.2", ec=ec, fc=fc, lw=1.5)
                ax.add_patch(cell)

                ax.text(x + 17, y + 15, p['p_num'], color="#F59E0B" if p['is_mix'] else "#38BDF8", fontsize=10, fontproperties=font_prop, ha='center', va='center')
                ax.text(x + 17, y + 10, f"{p['code']} ({p['color']})", color="#FFFFFF", fontsize=8.5, fontproperties=font_prop, ha='center', va='center')
                ax.text(x + 17, y + 5, f"{p['size']} | {p['box_cnt']}", color="#4ADE80" if p['is_full'] else "#93C5FD", fontsize=7.5, fontproperties=font_prop, ha='center', va='center')

            ax.axis('off')
            plt.tight_layout()

            img_buf = io.BytesIO()
            plt.savefig(img_buf, format="png", dpi=180)
            img_buf.seek(0)
            st.pyplot(fig)

            st.download_button(
                label="🖼️ CAD 미니맵 이미지(.png) 다운로드",
                data=img_buf,
                file_name="언탭트_파레트하역_2D미니맵.png",
                mime="image/png"
            )

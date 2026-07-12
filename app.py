import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import traceback
import csv
import time
from openpyxl.styles import Font, PatternFill, Border


# ──────────────────────────────────────────────────────────────────────────────
# Column index helper
# ──────────────────────────────────────────────────────────────────────────────

def ci(col: str) -> int:
    """Excel column letter(s) → 0-based int.  A=0, B=1, AA=26, BL=63 …"""
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


# ──────────────────────────────────────────────────────────────────────────────
# CSV 파일 종류 판별 (파일명이 아닌 컬럼 구조 기반)
# ──────────────────────────────────────────────────────────────────────────────

MEDIA_COLUMNS = [
    'Date', 'Year', 'Month', 'Week', 'Media', 'Campaign', 'Ad Group', 'ad',
    'Spend (gross)', 'Impressions', 'Click', 'Video Views', 'Reach',
    'AF Installs', 'AF Reinstalls', 'AF Revenue', 'AF Order Amount',
    'AF add_to_cart', 'AF Coupon', 'pb_order_now_delivery_quick',
    'pb_order_now_delivery_mart', 'pb_order_now_delivery_all',
    'pb_order_kurlynmart', 'pb_order_kurlynmart(Unique)',
    'pb_order_kurlynmart_revenue', 'pb_view_product_kurlynmart',
    'pb_view_product_kurlynmart(Unique)', 'kurly_view_home',
    'kurly_view_home(Unique)', 'first_order', 'temp_promo1', 'temp_promo2',
    'temp_event1', 'temp_event1(Unique)', 'OS', 'Campaign Theme',
    'Campaign Objective', 'Ad Product', 'Campaign Details', 'Targeting',
    'Gender', 'Age', 'Detailed Targeting', 'Creative Live Date',
    'Creative Format', 'Creative Type', 'Dimension', 'USP Category',
    'USP Brand', 'USP', '애드코드', 'Creative Full Name', '대구분',
    're-engagement', 'gross(net)', 'Campaign Theme(ADEF)',
    'Campaign Objective(ADEF)', 'Ad Product(ADEF)',
    'Detailed Targeting(ADEF)', 'USP Category(ADEF)', 'USP Brand(ADEF)',
    'USP(ADEF)', 'AF Order Count(Adef)', '집약형(Adef)', 'Install(SKAN)',
    '프로모션', '소재이미지', '소재카피', '소구점',
]

CATEGORY_EXTRA_COLUMNS = [
    '카테고리 구매 unique', '카테고리 구매 quantity', '카테고리 구매 price',
    '패션 unique', '패션 quantity', '패션 price',
    '뷰티 unique', '뷰티 quantity', '뷰티 price',
    '디지털가전 unique', '디지털가전 quantity', '디지털가전 price',
    '가구 unique', '가구 quantity', '가구 price',
    '키즈 unique', '키즈 quantity', '키즈 price',
    '식품 unique', '식품 quantity', '식품 price',
    '스포츠/레저 unique', '스포츠/레저 quantity', '스포츠/레저 price',
    '생활/건강 unique', '생활/건강 quantity', '생활/건강 price',
    '여가/생활편의 unique', '여가/생활편의 quantity', '여가/생활편의 price',
    '리빙 unique', '리빙 quantity', '리빙 price',
    '자동차/공구 unique', '자동차/공구 quantity', '자동차/공구 price',
    '펫 unique', '펫 quantity', '펫 price',
    'e-kam unique', 'e-kam quantity', 'e-kam price',
    '여가/도서(e쿠폰) unique', '여가/도서(e쿠폰) quantity', '여가/도서(e쿠폰) price',
    '노크잇 unique', '노크잇 quantity', '노크잇 price',
]

CATEGORY_COLUMNS  = MEDIA_COLUMNS + CATEGORY_EXTRA_COLUMNS
CATEGORY_MARKER   = '카테고리 구매 unique'  # 카테고리 파일에만 존재하는 컬럼


def detect_file_kind(df):
    """컬럼 구조로 파일 종류 판별. '카테고리 구매 unique' 컬럼 존재 여부가 기준."""
    if df is None:
        return None
    return 'category' if CATEGORY_MARKER in df.columns else 'media'


def _sniff_column_names(uploaded_file, encoding):
    """헤더 행과 첫 데이터 행의 실제 필드 수를 비교해 컬럼명 리스트를 만든다.
    데이터 행에 트레일링 콤마 등으로 헤더보다 필드가 더 많으면, 그 여분 필드에
    '_extra_N' 자리표시 이름을 붙여준다 — 이렇게 모든 필드에 명시적 이름을 지정해야
    pandas가 이름 없는 첫 열을 인덱스로 흡수하는 것을 막고 Date열이 항상 일반 컬럼으로
    읽히도록 보장할 수 있다."""
    uploaded_file.seek(0)
    header_line = uploaded_file.readline().decode(encoding)
    data_line   = uploaded_file.readline().decode(encoding)

    header_fields = next(csv.reader([header_line])) if header_line else []
    data_fields   = next(csv.reader([data_line]))   if data_line   else header_fields

    n_extra = max(0, len(data_fields) - len(header_fields))
    return header_fields + [f'_extra_{i}' for i in range(n_extra)]


def read_csv_robust(uploaded_file):
    """인코딩이 다른 CSV(UTF-8 / CP949 등)를 순차 시도하여 읽는다.
    헤더를 직접 읽어 컬럼명으로 명시 지정해서 읽기 때문에, 데이터 행의 필드 수가
    헤더보다 많은 경우(트레일링 콤마 등)에도 Date 등 첫 열이 인덱스로 잘못 흡수되지
    않고 항상 일반 컬럼으로 읽힌다."""
    for enc in ('utf-8-sig', 'utf-8', 'cp949'):
        try:
            names = _sniff_column_names(uploaded_file, enc)
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, header=0, names=names, encoding=enc)
            extra_cols = [c for c in df.columns if c.startswith('_extra_')]
            if extra_cols:
                df = df.drop(columns=extra_cols)
            return df
        except (UnicodeDecodeError, UnicodeError):
            continue
    uploaded_file.seek(0)
    return pd.read_csv(uploaded_file, header=0, index_col=False)  # 마지막 시도, 오류 그대로 노출


# ──────────────────────────────────────────────────────────────────────────────
# Column-index constants
# ──────────────────────────────────────────────────────────────────────────────

# Media file
M_DATE, M_CAMP, M_ADG, M_AD = ci('A'), ci('F'), ci('G'), ci('H')

# Category file
C_DATE, C_MEDIA, C_CAMP, C_ADG, C_AD = ci('A'), ci('E'), ci('F'), ci('G'), ci('H')
C_AV = ci('AV')                  # 47 – USP Category (사업부-연합 세부 분기용, 원본 그대로 유효)

# 사업부구분(인덱스 그룹 시트 기준) → 카테고리 그룹명. 그룹명 + ' unique'/' quantity'/' price'로
# CATEGORY_EXTRA_COLUMNS의 실제 열 이름을 만든다 (위치가 아닌 이름으로 조회 — BR열 등
# 위치 기반 매핑이 실제 컬럼 구조와 어긋났던 문제 재발 방지).
BIZ_TO_CATEGORY_GROUP = {
    '사업부-가구':          '가구',
    '사업부-그로서리-전체': '식품',
    '사업부-그로서리-별도': '식품',
    '사업부-리빙':          '리빙',
    '사업부-자동차공구':    '자동차/공구',
    '사업부-키즈':          '키즈',
    '사업부-펫':            '펫',
    '사업부-여가생활e쿠폰': '여가/도서(e쿠폰)',
}
USP_TO_CATEGORY_GROUP = {
    'LVG':     '리빙',
    'PET':     '펫',
    'CARTOOL': '자동차/공구',
    'KID':     '키즈',
}
GENERIC_CATEGORY_COLS = ('카테고리 구매 unique', '카테고리 구매 quantity', '카테고리 구매 price')


def category_group_cols(group_name):
    return (f'{group_name} unique', f'{group_name} quantity', f'{group_name} price')


# ──────────────────────────────────────────────────────────────────────────────
# 인덱스 파일 (그룹/소재 시트)
# ──────────────────────────────────────────────────────────────────────────────

INDEX_GROUP_COLUMNS    = ['사업부구분', 'Media', 'Campaign', 'Ad Group', '종료일']
INDEX_CREATIVE_COLUMNS = ['사업부구분', 'Media', 'Campaign', 'Ad Group', 'Ad',
                          '프로모션', '기여기간', '종료일']


def parse_index_file(uploaded_file):
    """인덱스 xlsx(그룹/소재 2개 시트)를 읽어 (group_df, creative_df)로 반환.
    시트명에 '그룹'/'소재'가 포함되면 이를 우선 사용하고, 없으면 시트 순서(1번째=그룹,
    2번째=소재)로 판별한다."""
    sheets = pd.read_excel(uploaded_file, sheet_name=None, header=0, engine='openpyxl')
    group_df = creative_df = None
    for name, df in sheets.items():
        if '그룹' in str(name):
            group_df = df
        elif '소재' in str(name):
            creative_df = df

    names = list(sheets.keys())
    if group_df is None and len(names) >= 1:
        group_df = sheets[names[0]]
    if creative_df is None and len(names) >= 2:
        creative_df = sheets[names[1]]

    return group_df, creative_df


INDEX_MISSING_MARK = '#인덱스추가'  # 인덱스에 없는 Campaign+Ad Group(+Ad) 표시값


def build_index_lookup(group_df):
    """그룹 시트 → {(Campaign, Ad Group): (사업부구분, 종료일 Timestamp)} 매핑."""
    lookup = {}
    if group_df is None or group_df.empty:
        return lookup
    if 'Campaign' not in group_df.columns or 'Ad Group' not in group_df.columns:
        return lookup

    biz_col = '사업부구분' if '사업부구분' in group_df.columns else None
    end_col = '종료일'    if '종료일'    in group_df.columns else None

    g = group_df.copy()
    if end_col:
        g[end_col] = parse_flexible_dates(g[end_col])

    for _, row in g.iterrows():
        key = (str(row['Campaign']).strip(), str(row['Ad Group']).strip())
        lookup[key] = (
            row[biz_col] if biz_col else np.nan,
            row[end_col] if end_col else pd.NaT,
        )
    return lookup


def build_creative_lookup(creative_df):
    """소재 시트 → {(Campaign, Ad Group, Ad): 종료일 Timestamp} 매핑."""
    lookup = {}
    if creative_df is None or creative_df.empty:
        return lookup
    if any(c not in creative_df.columns for c in ('Campaign', 'Ad Group', 'Ad')):
        return lookup

    end_col = '종료일' if '종료일' in creative_df.columns else None

    g = creative_df.copy()
    if end_col:
        g[end_col] = parse_flexible_dates(g[end_col])

    for _, row in g.iterrows():
        key = (str(row['Campaign']).strip(), str(row['Ad Group']).strip(), str(row['Ad']).strip())
        lookup[key] = row[end_col] if end_col else pd.NaT
    return lookup


def d7_status(row_date, end_date):
    """종료일+7일 기준 포함/제외 판정. 종료일을 알 수 없으면 기본값 '포함'."""
    if pd.isna(end_date) or pd.isna(row_date):
        return '포함'
    return '포함' if row_date <= end_date + pd.Timedelta(days=7) else '제외'


def add_index_columns(df, group_lookup, creative_lookup, camp_col, adg_col, ad_col, date_col):
    """그룹/소재 인덱스 매칭 결과를 추가한다 — 맨 앞 2열 '그룹_D7초과여부'/'소재_D7초과여부',
    맨 끝 1열 '사업부구분'. 원본 데이터 열 순서는 그대로 유지된다.
    그룹은 Campaign+Ad Group, 소재는 Campaign+Ad Group+Ad로 매칭한다. 행 삭제는 하지 않는다.
    - 그룹 미매칭 → 사업부구분/그룹_D7초과여부 모두 '#인덱스추가'
    - 소재 미매칭이지만 그룹은 매칭 → 소재_D7초과여부 = '그룹기준적용'
    - 그룹·소재 둘 다 미매칭 → 소재_D7초과여부도 '#인덱스추가'
    """
    out   = df.copy()
    dates = parse_flexible_dates(out[date_col])

    biz_list, group_status, creative_status = [], [], []
    for camp, adg, ad, d in zip(out[camp_col], out[adg_col], out[ad_col], dates):
        camp_s, adg_s, ad_s = str(camp).strip(), str(adg).strip(), str(ad).strip()

        gmatch = group_lookup.get((camp_s, adg_s))
        if gmatch is None:
            biz_list.append(INDEX_MISSING_MARK)
            group_status.append(INDEX_MISSING_MARK)
        else:
            biz_val, g_end = gmatch
            biz_list.append(biz_val if pd.notna(biz_val) else INDEX_MISSING_MARK)
            group_status.append(d7_status(d, g_end))

        c_end = creative_lookup.get((camp_s, adg_s, ad_s))
        if c_end is not None:
            creative_status.append(d7_status(d, c_end))
        elif gmatch is not None:
            creative_status.append('그룹기준적용')
        else:
            creative_status.append(INDEX_MISSING_MARK)

    out.insert(0, '소재_D7초과여부', creative_status)
    out.insert(0, '그룹_D7초과여부', group_status)
    out['사업부구분'] = biz_list
    return out


# ──────────────────────────────────────────────────────────────────────────────
# Utilities
# ──────────────────────────────────────────────────────────────────────────────

def safe_col(df, idx):
    return df.columns[idx] if idx < len(df.columns) else None


def row_val(row, idx, default=np.nan):
    return row.iloc[idx] if idx < len(row) else default


DATE_FALLBACK_FORMATS = ('%Y%m%d', '%Y/%m/%d', '%Y.%m.%d', '%Y-%m-%d')


def parse_flexible_dates(series):
    """열에 섞여 있는 여러 날짜 형식(2026-07-03, 20260703, 2026/07/03 등)을 최대한 인식한다.
    pandas.to_datetime은 한 열에 형식이 섞여 있으면 처음 인식한 형식만 유지하고 나머지를
    NaT로 만들어버리는 경우가 있어(형식별로는 개별적으로 모두 정상 파싱됨), 1차 파싱 후
    남은 값에 대해 형식을 하나씩 지정해 재시도한다."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    s = series.astype(str).str.strip()
    parsed  = pd.to_datetime(s, errors='coerce')
    missing = parsed.isna() & series.notna()

    for fmt in DATE_FALLBACK_FORMATS:
        if not missing.any():
            break
        attempt = pd.to_datetime(s[missing], format=fmt, errors='coerce')
        parsed.loc[missing] = attempt
        missing = parsed.isna() & series.notna()

    return parsed


# ──────────────────────────────────────────────────────────────────────────────
# Media processing
# ──────────────────────────────────────────────────────────────────────────────

def process_media(raw):
    """업로드된 파일의 모든 데이터를 날짜 범위 필터링 없이 그대로 가공한다."""
    df = raw.copy()
    date_col = df.columns[M_DATE]

    df[date_col] = parse_flexible_dates(df[date_col])
    n_before = len(df)
    df = df.dropna(subset=[date_col]).sort_values(date_col).reset_index(drop=True)

    if df.empty:
        st.warning(f"⚠️ 미디어 파일: A열(날짜)을 인식할 수 있는 행이 없습니다. "
                    f"날짜 형식을 확인해주세요. (원본 {n_before}행 중 0행 인식)")

    return df


# ──────────────────────────────────────────────────────────────────────────────
# Category processing
# ──────────────────────────────────────────────────────────────────────────────

def process_category_step1(raw):
    """1단계: 날짜 파싱/정렬 (컬럼 구조 검증은 호출부에서 처리). 원본 데이터 열은 그대로 유지.
    A열(날짜)을 찾을 수 없거나 인식 가능한 행이 하나도 없으면 ValueError를 낸다."""
    df = raw.copy()
    date_col = safe_col(df, C_DATE)
    if not date_col:
        raise ValueError("A열(날짜)를 찾을 수 없습니다.")

    df[date_col] = parse_flexible_dates(df[date_col])
    df = df.dropna(subset=[date_col]).sort_values(date_col).reset_index(drop=True)

    if df.empty:
        raise ValueError("A열(날짜)을 인식할 수 있는 행이 없습니다. 날짜 형식을 확인해주세요.")

    return df


def process_category_step2(df, group_lookup, creative_lookup):
    """2단계: 인덱스 매칭 — 사업부구분(그룹 시트 Campaign+Ad Group), 그룹_D7초과여부(그룹 시트
    종료일+7일), 소재_D7초과여부(소재 시트 Campaign+Ad Group+Ad)를 계산한다. 매칭되지 않는
    행도 삭제하지 않고 '#인덱스추가'로 표시한다.
    Returns (biz_list, group_status, creative_status) — 모두 len(df) 크기의 object 배열.
    """
    n = len(df)
    biz_list         = np.full(n, '', object)
    group_status     = np.full(n, '', object)
    creative_status  = np.full(n, '', object)

    for pos, (_, row) in enumerate(df.iterrows()):
        camp     = str(row_val(row, C_CAMP, '')).strip()
        adg      = str(row_val(row, C_ADG, '')).strip()
        ad       = str(row_val(row, C_AD, '')).strip()
        row_date = row_val(row, C_DATE)

        gmatch = group_lookup.get((camp, adg))
        if gmatch is None:
            biz_list[pos]     = INDEX_MISSING_MARK
            group_status[pos] = INDEX_MISSING_MARK
        else:
            biz_val, g_end = gmatch
            biz_list[pos]     = biz_val if pd.notna(biz_val) else INDEX_MISSING_MARK
            group_status[pos] = d7_status(row_date, g_end)

        c_end = creative_lookup.get((camp, adg, ad))
        if c_end is not None:
            creative_status[pos] = d7_status(row_date, c_end)
        elif gmatch is not None:
            creative_status[pos] = '그룹기준적용'
        else:
            creative_status[pos] = INDEX_MISSING_MARK

    return biz_list, group_status, creative_status


def process_category_step3(df, biz_list):
    """3단계: 사업부구분(2단계 결과) 기준 카테고리 값 매핑 — 카테고리 구매 unique/quantity/price
    3열의 값만 실제 카테고리 그룹 열(예: '가구 unique/quantity/price')의 값으로 치환하고,
    나머지 45열(패션·뷰티·...·노크잇)은 원본 값 그대로 둔다. 매핑 실패(사업부구분 미매핑,
    사업부-연합의 USP 미매핑, 인덱스 자체 미매칭)는 원본 총계 값을 그대로 두고 수기확인
    플래그만 세운다.
    Returns (df, manual_bool_array).
    """
    n      = len(df)
    manual = np.zeros(n, bool)

    generic_u, generic_q, generic_p = GENERIC_CATEGORY_COLS
    # object dtype으로 복사 — pandas 3.x의 엄격한 dtype(예: string[pyarrow])에 다른 타입 값을
    # 대입할 때 발생하는 TypeError를 피하기 위함(앞서 4~7일치 0-처리에서 겪은 문제와 동일한 원인).
    out_u = df[generic_u].astype(object).copy() if generic_u in df.columns else pd.Series([np.nan] * n, dtype=object)
    out_q = df[generic_q].astype(object).copy() if generic_q in df.columns else pd.Series([np.nan] * n, dtype=object)
    out_p = df[generic_p].astype(object).copy() if generic_p in df.columns else pd.Series([np.nan] * n, dtype=object)

    for pos, (_, row) in enumerate(df.iterrows()):
        biz = biz_list[pos]
        biz_str = str(biz).strip() if biz not in ('', None) and biz != INDEX_MISSING_MARK else ''

        group_name = None
        if biz_str == '사업부-연합':
            usp = str(row_val(row, C_AV, '')).strip().upper()
            group_name = USP_TO_CATEGORY_GROUP.get(usp)
            if group_name is None:
                manual[pos] = True
        elif biz_str:
            group_name = BIZ_TO_CATEGORY_GROUP.get(biz_str)
            if group_name is None:
                manual[pos] = True
        else:
            manual[pos] = True  # 인덱스에 없는 Campaign+Ad Group

        if group_name:
            u_name, q_name, p_name = category_group_cols(group_name)
            out_u.iloc[pos] = row[u_name] if u_name in df.columns else np.nan
            out_q.iloc[pos] = row[q_name] if q_name in df.columns else np.nan
            out_p.iloc[pos] = row[p_name] if p_name in df.columns else np.nan
        # group_name이 없으면(미매칭) 원본 '카테고리 구매' 값(총계 폴백) 그대로 유지

    if generic_u in df.columns:
        df[generic_u] = out_u
        df[generic_q] = out_q
        df[generic_p] = out_p

    return df, manual


def process_category_finalize(df, biz_list, group_status, creative_status):
    """열 순서 재배치 — 소구점 다음에 사업부구분/여백(빈값)/피드구분(빈값) 삽입,
    맨 앞에 그룹_D7초과여부/소재_D7초과여부 삽입. 원본 데이터 열 순서는 그대로 유지."""
    df = df.copy()
    if '소구점' in df.columns:
        pos_after = df.columns.get_loc('소구점') + 1
        df.insert(pos_after, '사업부구분', biz_list)
        df.insert(pos_after + 1, '여백', '')
        df.insert(pos_after + 2, '피드구분', '')
    else:
        df['사업부구분'] = biz_list
        df['여백'] = ''
        df['피드구분'] = ''

    df.insert(0, '소재_D7초과여부', creative_status)
    df.insert(0, '그룹_D7초과여부', group_status)
    return df


# ──────────────────────────────────────────────────────────────────────────────
# 화면 표시 전 Arrow 직렬화 안전화
# ──────────────────────────────────────────────────────────────────────────────

# ──────────────────────────────────────────────────────────────────────────────
# Excel export builder
# ──────────────────────────────────────────────────────────────────────────────

def _strip_header_style(writer, sheet_name):
    """헤더(1행) 서식을 일반 텍스트로 초기화 — 색상/볼드 등 어떤 스타일도 없이 출력한다."""
    ws = writer.sheets[sheet_name]
    for cell in next(ws.iter_rows(min_row=1, max_row=1)):
        cell.font = Font()
        cell.fill = PatternFill()
        cell.border = Border()


def build_media_excel(media_df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        media_df.to_excel(writer, sheet_name='미디어_가공', index=False)
        _strip_header_style(writer, '미디어_가공')
    buf.seek(0)
    return buf.getvalue()


def build_category_excel(cat_df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        cat_df.to_excel(writer, sheet_name='카테고리_전체', index=False)
        _strip_header_style(writer, '카테고리_전체')
    buf.seek(0)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────────────────────────────────────────

def _init_session_state():
    defaults = {
        'index_group_df':     None,
        'index_creative_df':  None,
        'index_filename':     None,
        'index_uploaded_at':  None,
        'media_log':          None,
        'cat_log':            None,
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


# ──────────────────────────────────────────────────────────────────────────────
# 가공 내역(단계별 진행 상태) 패널
# ──────────────────────────────────────────────────────────────────────────────

MEDIA_STEP_LABELS = [
    'CSV 파일 읽기 및 컬럼 구조 검증',
    '인덱스 매칭 — 사업부구분 / D7 판정',
    '엑셀 파일 생성',
]
CATEGORY_STEP_LABELS = [
    'CSV 파일 읽기 및 컬럼 구조 검증',
    '인덱스 매칭 — 사업부구분 / D7 판정',
    '카테고리 값 매핑 (사업부구분 기준)',
    '엑셀 파일 생성',
]

STATUS_BADGE = {
    'done':    ('완료',   'green', '✅'),
    'error':   ('오류',   'red',   '❌'),
    'pending': ('대기',   'gray',  '⚪'),
    'running': ('진행중', 'blue',  '🔵'),
}


def _new_run_log(step_labels):
    return {
        'steps':       [{'label': s, 'status': 'pending', 'detail': None} for s in step_labels],
        'rows':        None,
        'excel_bytes': None,
        'excel_fname': None,
    }


def render_run_log(icon, title, log, key_prefix):
    """가공 내역 패널: 제목 행(제목 + 진행률 + 다운로드 버튼) + 단계별 상태 리스트(라인 구분,
    배경 없음). 다운로드 버튼은 가공 완료 전엔 회색 비활성화, 완료 후엔 초록색 활성화로
    표시한다(Streamlit 기본 type="primary"는 앱 테마색(빨강)이라 st.container(key=)로 감싸서
    CSS를 직접 덮어씌운다). 전체적으로 폰트를 줄이고 여백을 좁혀 compact하게 구성한다.

    같은 스크립트 실행 안에서 이 함수가 여러 번 호출되면(단계 진행에 따라 실시간으로
    다시 그릴 때) 컨테이너 key를 고정값으로 두면 Streamlit이 첫 호출 내용에서 갱신을
    멈춰버린다(실측으로 확인된 동작) — 그래서 호출마다 새로운 key를 사용한다.

    이 일련번호는 반드시 st.session_state에 저장해야 한다 — log 딕셔너리에 저장하면
    (예: log['_seq']) 새 가공을 시작할 때마다 _new_run_log()가 완전히 새 log를 만들면서
    번호가 다시 0부터 시작해, 이전 실행 때와 동일한 key가 재사용된다. 그 결과 이전 실행에서
    렌더링된 단계별 상세 텍스트(예: 잘못된 파일 형식 오류 메시지)가 새 실행에서 그 텍스트를
    다시 쓰지 않아도(caption을 호출 안 해도) 화면에 그대로 눌어붙어 있는 버그가 실측으로
    확인됐다. session_state는 실행 간에도 유지되므로 여기 저장해야 매 호출이 진짜 유일한
    key를 갖는다."""
    steps       = log['steps']
    total       = len(steps)
    done_n      = sum(1 for s in steps if s['status'] == 'done')
    has_error   = any(s['status'] == 'error' for s in steps)
    progress    = done_n / total if total else 0.0
    can_download = bool(log['excel_bytes']) and not has_error

    seq_state_key = f"_{key_prefix}_render_seq"
    seq = st.session_state.get(seq_state_key, 0) + 1
    st.session_state[seq_state_key] = seq

    panel_key = f"{key_prefix}_panel_{seq}"
    dl_key    = f"{key_prefix}_dl_wrap_{seq}"
    step_rules = "\n".join(
        f".st-key-{key_prefix}_step_{i}_{seq} {{ border-bottom: 1px solid rgba(49,51,63,0.08); "
        f"padding: 0.3rem 0; margin-bottom: 0; }}"
        for i in range(total)
    )

    st.markdown(f"""
        <style>
        .st-key-{panel_key} {{ font-size: 0.85rem; }}
        .st-key-{panel_key} h4 {{ font-size: 1rem; margin: 0; padding: 0; }}
        .st-key-{panel_key} [data-testid="stCaptionContainer"] p {{ font-size: 0.78rem; }}
        .st-key-{panel_key} [data-testid="stProgress"] {{ margin-top: 0.4rem; }}
        .st-key-{panel_key} [data-testid="stProgressBar"] > div > div {{ height: 4px !important; }}
        .st-key-{panel_key} span[data-testid="stBadge"] {{
            font-size: 0.72rem !important;
            padding: 0.05rem 0.5rem !important;
        }}
        {step_rules}
        .st-key-{dl_key} button {{
            padding: 0.25rem 0.75rem;
            box-shadow: none;
        }}
        .st-key-{dl_key} button:disabled {{
            background-color: #f1f5f9 !important;
            color: #b0b8c1 !important;
            border-color: #f1f5f9 !important;
        }}
        .st-key-{dl_key} button:not(:disabled) {{
            background-color: #86efac !important;
            color: #14532d !important;
            border-color: #86efac !important;
        }}
        </style>
    """, unsafe_allow_html=True)

    with st.container(key=panel_key):
        head_l, head_r = st.columns([2.2, 2.8], vertical_alignment="center")
        with head_l:
            st.markdown(f"#### {icon} {title}")
        with head_r:
            c2, c3 = st.columns([2, 1.2], vertical_alignment="center")
            with c2:
                st.progress(progress)
            with c3:
                with st.container(key=dl_key):
                    st.download_button(
                        "📥 다운로드",
                        data=log['excel_bytes'] or b'',
                        file_name=log['excel_fname'] or 'download.xlsx',
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        disabled=not can_download,
                        use_container_width=True,
                        key=f"{key_prefix}_dl_btn_{seq}",
                    )

        for i, s in enumerate(steps):
            label, color, badge_icon = STATUS_BADGE[s['status']]
            with st.container(key=f"{key_prefix}_step_{i}_{seq}"):
                col1, col2, col3 = st.columns([0.4, 3.6, 1], vertical_alignment="center")
                with col1:
                    st.markdown(badge_icon)
                with col2:
                    st.write(s['label'])
                    # detail 유무와 무관하게 매번 st.caption을 호출한다(빈 문자열이라도) — 조건부로
                    # 호출을 건너뛰면 다음 실행에서 위젯 호출 순서/개수가 달라져 Streamlit이 이전
                    # 실행의 캡션 내용을 새 컨테이너에 그대로 흘려보내는 현상이 실측으로 확인됨.
                    detail_text = s['detail'] or ""
                    st.caption(f":red[{detail_text}]" if s['status'] == 'error' and detail_text else detail_text)
                with col3:
                    st.badge(label, color=color)


def main():
    st.set_page_config(page_title="NPS Report 가공기", layout="wide", page_icon="📊")
    st.title("📊 NPS Report 데이터 가공기")
    st.caption("미디어·카테고리 CSV 로우 데이터를 업로드하면 RD 시트 형식으로 자동 가공합니다.")

    _init_session_state()

    # ── 파일 업로드 (섹션별로 제목 행 오른쪽에 버튼, 업로더는 그 아래 전체 폭)
    # 제목 + 버튼 행 — 중첩 컬럼이 아닌 하나의 행(6칸)으로 구성해야 좁은 폭에서도
    # 같은 줄을 유지하며, vertical_alignment="center"로 제목 텍스트와 버튼을 수직 중앙 정렬한다.
    ht1, hb1, ht2, hb2, ht3, hb3 = st.columns(
        [2.3, 0.8, 2.3, 0.8, 2.3, 0.8], vertical_alignment="center"
    )
    with ht1:
        st.markdown("### 📁 미디어 파일")
    with hb1:
        run_media = st.button("가공", key='media_btn', type="primary", use_container_width=True)

    with ht2:
        st.markdown("### 📁 카테고리 파일")
    with hb2:
        run_cat = st.button("가공", key='cat_btn', type="primary", use_container_width=True)

    with ht3:
        st.markdown("### 🗂️ 인덱스 파일")
    with hb3:
        run_index = st.button("업로드", key='index_btn', type="primary", use_container_width=True)

    # 업로더 행 — 버튼과 별도 행이므로 섹션 전체 폭을 그대로 채운다.
    u1, u2, u3 = st.columns(3)
    with u1:
        media_file = st.file_uploader("미디어 csv 업로드", type=['csv'], key='mf',
                                       label_visibility='collapsed')
    with u2:
        cat_file = st.file_uploader("카테고리 csv 업로드", type=['csv'], key='cf',
                                     label_visibility='collapsed')
    with u3:
        index_file = st.file_uploader("인덱스 xlsx 업로드", type=['xlsx', 'xls'], key='ixf',
                                       label_visibility='collapsed')

    # 상태 메시지 행
    s1, s2, s3 = st.columns(3)
    with s1:
        if media_file:
            st.success(f"✅ {media_file.name}")
    with s2:
        if cat_file:
            st.success(f"✅ {cat_file.name}")
    with s3:
        if run_index:
            if not index_file:
                st.warning("인덱스 파일을 먼저 선택해주세요.")
            else:
                try:
                    group_df, creative_df = parse_index_file(index_file)
                    missing = [c for c in INDEX_GROUP_COLUMNS if c not in group_df.columns]
                    if missing:
                        st.error(f"인덱스 그룹 시트에 필요한 열이 없습니다: {', '.join(missing)}")
                    else:
                        st.session_state['index_group_df']    = group_df
                        st.session_state['index_creative_df'] = creative_df
                        st.session_state['index_filename']    = index_file.name
                        st.session_state['index_uploaded_at'] = pd.Timestamp.now()
                except Exception as e:
                    st.error(f"인덱스 파일 읽기 오류: {e}")

        if st.session_state['index_group_df'] is not None:
            up_at = st.session_state['index_uploaded_at']
            st.success(f"✅ 현재 적용 중인 인덱스: **{st.session_state['index_filename']}** "
                       f"({up_at.strftime('%Y-%m-%d %H:%M')})")
            g_n = len(st.session_state['index_group_df'])
            c_n = len(st.session_state['index_creative_df']) \
                if st.session_state['index_creative_df'] is not None else 0
            st.caption(f"그룹 시트 {g_n}건 · 소재 시트 {c_n}건")
        else:
            st.info("적용된 인덱스가 없습니다. 파일 선택 후 업로드 버튼을 눌러주세요.")

        st.caption("그룹 시트: 사업부구분 · Media · Campaign · Ad Group · 종료일\n\n"
                   "업로드 후에는 다시 올리지 않아도 계속 적용됩니다.")

    st.divider()

    group_lookup    = build_index_lookup(st.session_state['index_group_df'])
    creative_lookup = build_creative_lookup(st.session_state['index_creative_df'])

    if st.session_state['media_log'] is None and st.session_state['cat_log'] is None \
            and not run_media and not run_cat:
        st.info("파일을 업로드한 후 각 섹션의 버튼을 눌러주세요.")

    # st.empty()로 자리를 잡기 전에 먼저 보여줘야 하는 안내는 여기서 처리한다 — placeholder를
    # 만든 뒤에 일반 st.warning()을 호출하면 placeholder가 이미 차지한 자리보다 아래쪽에
    # 그려져 버려서, 가공 내역 패널보다 밑에 경고가 나오는 순서 뒤바뀜이 생긴다.
    if run_cat and cat_file and not group_lookup:
        st.warning("⚠️ 적용된 인덱스가 없어 사업부구분을 판별할 수 없습니다. "
                   "전체 행이 '#인덱스추가'로 표시되고 카테고리 구매 unique/quantity/price는 "
                   "원본 총계 값 그대로 유지됩니다. 인덱스 파일을 먼저 업로드해주세요.")

    # 가공 내역 영역 — 버튼을 누른 즉시 자리부터 잡아두고(placeholder), 단계가 진행될 때마다
    # 그 자리를 다시 그려서 완료 전에도 실시간으로 진행 상황이 보이도록 한다.
    media_slot = st.empty()
    cat_slot   = st.empty()

    def _refresh_media(log, pause=0.15):
        with media_slot.container():
            render_run_log("📊", "미디어 가공 내역", log, key_prefix="media")
        if pause:
            time.sleep(pause)

    def _refresh_cat(log, pause=0.15):
        with cat_slot.container():
            if st.session_state['media_log'] is not None:
                st.divider()
            render_run_log("🛒", "카테고리 가공 내역", log, key_prefix="cat")
        if pause:
            time.sleep(pause)

    # ────────────────────────────── 미디어 가공 ──────────────────────────────
    if run_media:
        if not media_file:
            st.warning("미디어 파일을 업로드해주세요.")
        else:
            st.session_state['media_log'] = None  # 이전 가공(오류 포함) 내역 즉시 초기화
            log = _new_run_log(MEDIA_STEP_LABELS)
            _refresh_media(log)  # 버튼 클릭 즉시 전체 대기 상태로 먼저 표시

            # 1단계: CSV 파일 읽기 및 컬럼 구조 검증
            m_df = None
            log['steps'][0]['status'] = 'running'
            _refresh_media(log)
            try:
                media_raw = read_csv_robust(media_file)
                if detect_file_kind(media_raw) == 'category':
                    raise ValueError(
                        f"'{media_file.name}'은(는) 카테고리 파일 컬럼 구조로 보입니다. "
                        "카테고리 파일 업로드란에 올린 뒤 카테고리 가공 버튼을 사용해주세요."
                    )
                col_note = None
                if media_raw.shape[1] != len(MEDIA_COLUMNS):
                    col_note = (f"컬럼 수({media_raw.shape[1]})가 예상({len(MEDIA_COLUMNS)}개)과 "
                                "다릅니다. 컬럼 구조를 확인해주세요.")
                m_df = process_media(media_raw)
                if m_df.empty:
                    raise ValueError("A열(날짜)을 인식할 수 있는 행이 없습니다. 날짜 형식을 확인해주세요.")
                log['steps'][0]['status'] = 'done'
                log['steps'][0]['detail'] = col_note
            except Exception as e:
                log['steps'][0]['status'] = 'error'
                log['steps'][0]['detail'] = str(e)
                m_df = None
            _refresh_media(log)

            # 2단계: 인덱스 매칭 — 사업부구분 / D7 판정
            if log['steps'][0]['status'] == 'done':
                log['steps'][1]['status'] = 'running'
                _refresh_media(log)
                try:
                    date_col = m_df.columns[M_DATE]
                    camp_col = m_df.columns[M_CAMP]
                    adg_col  = m_df.columns[M_ADG]
                    ad_col   = m_df.columns[M_AD]
                    m_df = add_index_columns(
                        m_df, group_lookup, creative_lookup,
                        camp_col, adg_col, ad_col, date_col
                    )
                    n_missing = int((m_df['사업부구분'] == INDEX_MISSING_MARK).sum())
                    log['steps'][1]['status'] = 'done'
                    log['steps'][1]['detail'] = f"#인덱스추가 {n_missing}건" if n_missing else "전체 매칭 완료"
                    log['rows'] = len(m_df)
                except Exception:
                    log['steps'][1]['status'] = 'error'
                    log['steps'][1]['detail'] = traceback.format_exc()
                _refresh_media(log)

            # 3단계: 엑셀 파일 생성
            if log['steps'][1]['status'] == 'done':
                log['steps'][2]['status'] = 'running'
                _refresh_media(log)
                try:
                    log['excel_bytes'] = build_media_excel(m_df)
                    log['excel_fname'] = f"미디어_가공_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    log['steps'][2]['status'] = 'done'
                except Exception:
                    log['steps'][2]['status'] = 'error'
                    log['steps'][2]['detail'] = traceback.format_exc()
                _refresh_media(log, pause=0)

            st.session_state['media_log'] = log
    elif st.session_state['media_log'] is not None:
        _refresh_media(st.session_state['media_log'], pause=0)

    # ────────────────────────────── 카테고리 가공 ────────────────────────────
    if run_cat:
        if not cat_file:
            st.warning("카테고리 파일을 업로드해주세요.")
        else:
            st.session_state['cat_log'] = None  # 이전 가공(오류 포함) 내역 즉시 초기화
            log = _new_run_log(CATEGORY_STEP_LABELS)
            _refresh_cat(log)  # 버튼 클릭 즉시 전체 대기 상태로 먼저 표시

            # 1단계: CSV 파일 읽기 및 컬럼 구조 검증
            df1 = None
            log['steps'][0]['status'] = 'running'
            _refresh_cat(log)
            try:
                cat_raw = read_csv_robust(cat_file)
                if detect_file_kind(cat_raw) == 'media':
                    raise ValueError(
                        f"'{cat_file.name}'은(는) 미디어 파일 컬럼 구조로 보입니다. "
                        "미디어 파일 업로드란에 올린 뒤 미디어 가공 버튼을 사용해주세요."
                    )
                col_note = None
                if cat_raw.shape[1] != len(CATEGORY_COLUMNS):
                    col_note = (f"컬럼 수({cat_raw.shape[1]})가 예상({len(CATEGORY_COLUMNS)}개)과 "
                                "다릅니다. 컬럼 구조를 확인해주세요.")
                df1 = process_category_step1(cat_raw)
                log['steps'][0]['status'] = 'done'
                log['steps'][0]['detail'] = col_note
            except Exception as e:
                log['steps'][0]['status'] = 'error'
                log['steps'][0]['detail'] = str(e)
                df1 = None
            _refresh_cat(log)

            # 2단계: 인덱스 매칭 — 사업부구분 / D7 판정
            biz_list = group_status = creative_status = None
            if log['steps'][0]['status'] == 'done':
                log['steps'][1]['status'] = 'running'
                _refresh_cat(log)
                try:
                    biz_list, group_status, creative_status = process_category_step2(
                        df1, group_lookup, creative_lookup
                    )
                    n_missing = int((biz_list == INDEX_MISSING_MARK).sum())
                    log['steps'][1]['status'] = 'done'
                    log['steps'][1]['detail'] = f"#인덱스추가 {n_missing}건" if n_missing else "전체 매칭 완료"
                except Exception:
                    log['steps'][1]['status'] = 'error'
                    log['steps'][1]['detail'] = traceback.format_exc()
                _refresh_cat(log)

            # 3단계: 카테고리 값 매핑 (사업부구분 기준)
            df2 = None
            if log['steps'][1]['status'] == 'done':
                log['steps'][2]['status'] = 'running'
                _refresh_cat(log)
                try:
                    df2, manual = process_category_step3(df1, biz_list)
                    n_manual = int(manual.sum())
                    log['steps'][2]['status'] = 'done'
                    log['steps'][2]['detail'] = f"수기확인 필요 {n_manual}건" if n_manual else "전체 매핑 완료"
                except Exception:
                    log['steps'][2]['status'] = 'error'
                    log['steps'][2]['detail'] = traceback.format_exc()
                _refresh_cat(log)

            # 4단계: 엑셀 파일 생성
            if log['steps'][2]['status'] == 'done':
                log['steps'][3]['status'] = 'running'
                _refresh_cat(log)
                try:
                    final_df = process_category_finalize(df2, biz_list, group_status, creative_status)
                    log['rows'] = len(final_df)
                    log['excel_bytes'] = build_category_excel(final_df)
                    log['excel_fname'] = f"카테고리_가공_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    log['steps'][3]['status'] = 'done'
                except Exception:
                    log['steps'][3]['status'] = 'error'
                    log['steps'][3]['detail'] = traceback.format_exc()
                _refresh_cat(log, pause=0)

            st.session_state['cat_log'] = log
    elif st.session_state['cat_log'] is not None:
        _refresh_cat(st.session_state['cat_log'], pause=0)


if __name__ == "__main__":
    main()

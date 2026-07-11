import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import traceback


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


def read_csv_robust(uploaded_file):
    """인코딩이 다른 CSV(UTF-8 / CP949 등)를 순차 시도하여 읽는다."""
    for enc in ('utf-8-sig', 'utf-8', 'cp949'):
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, header=0, encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    uploaded_file.seek(0)
    return pd.read_csv(uploaded_file, header=0)  # 마지막 시도, 오류 그대로 노출


# ──────────────────────────────────────────────────────────────────────────────
# Column-index constants
# ──────────────────────────────────────────────────────────────────────────────

# Media file
M_DATE,  M_CAMP,  M_ADG  = ci('A'), ci('F'), ci('G')
M_NUM_ST, M_BL            = ci('I'), ci('BL')   # numeric block: I(8)..BL(63)
M_BL_NAME = MEDIA_COLUMNS[M_BL]                 # '집약형(Adef)' — 열 삽입 후에도 이름으로 참조

# Category file
C_DATE, C_MEDIA, C_CAMP, C_ADG = ci('A'), ci('E'), ci('F'), ci('G')
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
        g[end_col] = pd.to_datetime(g[end_col], errors='coerce')

    for _, row in g.iterrows():
        key = (str(row['Campaign']).strip(), str(row['Ad Group']).strip())
        lookup[key] = (
            row[biz_col] if biz_col else np.nan,
            row[end_col] if end_col else pd.NaT,
        )
    return lookup


def add_index_columns(df, lookup, camp_col, adg_col, date_col):
    """인덱스 매칭 결과를 맨 앞 두 열('사업부구분(인덱스)', 'D7_상태')로 추가.
    Campaign + Ad Group 매칭. 행 삭제는 하지 않고 상태만 표시(포함/제외)."""
    out   = df.copy()
    dates = pd.to_datetime(out[date_col], errors='coerce')

    biz_list, status_list = [], []
    for camp, adg, d in zip(out[camp_col], out[adg_col], dates):
        match = lookup.get((str(camp).strip(), str(adg).strip()))
        if match is None:
            biz_list.append(np.nan)
            status_list.append('포함')
            continue
        biz, end_date = match
        biz_list.append(biz)
        if pd.isna(end_date) or pd.isna(d):
            status_list.append('포함')
        else:
            status_list.append('포함' if d <= end_date + pd.Timedelta(days=7) else '제외')

    out.insert(0, 'D7_상태', status_list)
    out.insert(0, '사업부구분(인덱스)', biz_list)
    return out


# ──────────────────────────────────────────────────────────────────────────────
# Utilities
# ──────────────────────────────────────────────────────────────────────────────

def safe_col(df, idx):
    return df.columns[idx] if idx < len(df.columns) else None


def row_val(row, idx, default=np.nan):
    return row.iloc[idx] if idx < len(row) else default


# ──────────────────────────────────────────────────────────────────────────────
# Media processing
# ──────────────────────────────────────────────────────────────────────────────

def process_media(raw, yesterday):
    """
    Returns (processed_df, is_d4_7_bool_array).
    • 1~3일치: 전체 열 그대로
    • 4~7일치: I열~BK열 → 0, BL열 원본 유지
    """
    if raw.shape[1] <= M_BL:
        st.error(f"미디어 파일에 BL열({M_BL + 1}번째 열) 이상이 필요합니다. "
                 f"현재 열 수: {raw.shape[1]}")
        return None, None

    df = raw.copy()
    date_col = df.columns[M_DATE]

    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col]).sort_values(date_col).reset_index(drop=True)

    d7_start = yesterday - pd.Timedelta(days=6)
    d3_start = yesterday - pd.Timedelta(days=2)  # 3일치 중 가장 오래된 날

    df = df[(df[date_col] >= d7_start) & (df[date_col] <= yesterday)].copy()
    df = df.reset_index(drop=True)

    is_d4_7 = (df[date_col] < d3_start).values  # bool array aligned with df

    # Zero out I(8)..BK(62) for day-4-to-7 rows; BL(63) kept
    for col_i in range(M_NUM_ST, M_BL):
        if col_i < df.shape[1]:
            df.iloc[is_d4_7, col_i] = 0

    return df, is_d4_7


# ──────────────────────────────────────────────────────────────────────────────
# Category processing
# ──────────────────────────────────────────────────────────────────────────────

def process_category(raw, yesterday, lookup):
    """
    Returns (result_df, needs_manual_bool_array).
    result_df columns: 사업부구분(인덱스), D7_상태, 날짜, Media, Campaign, Ad Group,
                       USP_Category, 카테고리_unique, 카테고리_qty, 카테고리_price, 수기확인필요

    사업부구분은 카테고리 파일 자체에는 없고(BR열은 '카테고리 구매 unique' 수치 데이터)
    인덱스 그룹 시트(Campaign+Ad Group 매칭)에서만 가져온다. 이 값으로 카테고리_unique/qty/price를
    어느 카테고리 그룹 열(이름 기준)에서 가져올지 결정한다.
    """
    df = raw.copy()
    date_col = safe_col(df, C_DATE)
    if not date_col:
        st.error("카테고리 파일: A열(날짜)를 찾을 수 없습니다.")
        return pd.DataFrame(), np.array([], dtype=bool)

    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col])
    df = df[df[date_col].dt.normalize() == yesterday].copy().reset_index(drop=True)

    if df.empty:
        st.warning(f"카테고리 파일: 전일자({yesterday.date()}) 데이터가 없습니다.")
        return pd.DataFrame(), np.array([], dtype=bool)

    camp_col = safe_col(df, C_CAMP)
    adg_col  = safe_col(df, C_ADG)

    n        = len(df)
    out_u    = np.full(n, np.nan, object)
    out_q    = np.full(n, np.nan, object)
    out_p    = np.full(n, np.nan, object)
    manual   = np.zeros(n, bool)
    biz_list = np.full(n, np.nan, object)
    d7_list  = ['포함'] * n

    for pos, (_, row) in enumerate(df.iterrows()):
        camp = str(row_val(row, C_CAMP, '')).strip() if camp_col else ''
        adg  = str(row_val(row, C_ADG, '')).strip()  if adg_col  else ''
        match = lookup.get((camp, adg))

        biz      = match[0] if match else np.nan
        end_date = match[1] if match else pd.NaT
        biz_list[pos] = biz
        biz_str  = str(biz).strip() if pd.notna(biz) else ''

        row_date = row_val(row, C_DATE)
        if pd.notna(end_date) and pd.notna(row_date) and row_date > end_date + pd.Timedelta(days=7):
            d7_list[pos] = '제외'

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

        u_name, q_name, p_name = category_group_cols(group_name) if group_name else GENERIC_CATEGORY_COLS
        out_u[pos] = row[u_name] if u_name in df.columns else np.nan
        out_q[pos] = row[q_name] if q_name in df.columns else np.nan
        out_p[pos] = row[p_name] if p_name in df.columns else np.nan

    result = pd.DataFrame({
        '사업부구분(인덱스)': biz_list,
        'D7_상태':           d7_list,
        '날짜':             df.iloc[:, C_DATE].dt.date      if C_DATE  < df.shape[1] else '',
        'Media':            df.iloc[:, C_MEDIA]             if C_MEDIA < df.shape[1] else '',
        'Campaign':         df.iloc[:, C_CAMP]              if C_CAMP  < df.shape[1] else '',
        'Ad Group':         df.iloc[:, C_ADG]               if C_ADG   < df.shape[1] else '',
        'USP_Category':     df.iloc[:, C_AV]                if C_AV    < df.shape[1] else '',
        '카테고리_unique':  out_u,
        '카테고리_qty':     out_q,
        '카테고리_price':   out_p,
        '수기확인필요':     manual,
    })
    return result, manual


# ──────────────────────────────────────────────────────────────────────────────
# Row-highlight helper
# ──────────────────────────────────────────────────────────────────────────────

def highlight_rows(df, mask, color):
    """Return pandas Styler with rows highlighted where mask[i] is True."""
    def _fn(x):
        bg = pd.DataFrame('', index=x.index, columns=x.columns)
        if mask is not None:
            for i, flag in enumerate(mask):
                if flag and i < len(x):
                    bg.iloc[i] = f'background-color: {color}'
        return bg
    return df.style.apply(_fn, axis=None)


# ──────────────────────────────────────────────────────────────────────────────
# Excel export builder
# ──────────────────────────────────────────────────────────────────────────────

def build_excel(media_df, cat_df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        if media_df is not None and not media_df.empty:
            media_df.to_excel(writer, sheet_name='미디어_가공', index=False)

        if cat_df is not None and not cat_df.empty:
            rd_cols = ['날짜', 'Media', 'Campaign', 'Ad Group',
                       '카테고리_unique', '카테고리_qty', '카테고리_price']
            rd = cat_df[[c for c in rd_cols if c in cat_df.columns]]
            rd.to_excel(writer, sheet_name='카테고리_RD붙여넣기', index=False)
            cat_df.to_excel(writer, sheet_name='카테고리_전체', index=False)

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
        'index_sig':          None,
        'media_df':           None,
        'media_d4_7':         None,
        'cat_df':             None,
        'cat_manual':         None,
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


def main():
    st.set_page_config(page_title="NPS Report 가공기", layout="wide", page_icon="📊")
    st.title("📊 NPS Report 데이터 가공기")
    st.caption("미디어·카테고리 CSV 로우 데이터를 업로드하면 RD 시트 형식으로 자동 가공합니다.")

    _init_session_state()

    # ── 날짜 설정
    with st.expander("⚙️ 기준 날짜 설정 (기본: 자동 전일자)", expanded=False):
        use_custom  = st.checkbox("날짜 직접 지정")
        custom_date = st.date_input(
            "기준일 (이 날짜를 '전일자'로 간주)",
            value=(pd.Timestamp.now() - pd.Timedelta(days=1)).date()
        )
    yesterday = pd.Timestamp(custom_date) if use_custom \
        else pd.Timestamp.now().normalize() - pd.Timedelta(days=1)
    st.caption(f"현재 기준 전일자: **{yesterday.date()}**  |  "
               f"최근 3일: {(yesterday - pd.Timedelta(days=2)).date()} ~ {yesterday.date()}  |  "
               f"4~7일: {(yesterday - pd.Timedelta(days=6)).date()} ~ {(yesterday - pd.Timedelta(days=3)).date()}")

    st.divider()

    # ── 파일 업로드
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("### 📁 미디어 파일")
        media_file = st.file_uploader("미디어 csv 업로드", type=['csv'], key='mf')
        if media_file:
            st.success(f"✅ {media_file.name}")

    with c2:
        st.markdown("### 📁 카테고리 파일")
        cat_file = st.file_uploader("카테고리 csv 업로드", type=['csv'], key='cf')
        if cat_file:
            st.success(f"✅ {cat_file.name}")

    with c3:
        st.markdown("### 🗂️ 인덱스 파일")
        index_file = st.file_uploader("인덱스 xlsx 업로드", type=['xlsx', 'xls'], key='ixf')

        if index_file is not None:
            sig = (index_file.name, index_file.size)
            if st.session_state['index_sig'] != sig:
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
                        st.session_state['index_sig']         = sig
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
            st.info("적용된 인덱스가 없습니다. 업로드하면 자동 저장됩니다.")

        st.caption("그룹 시트: 사업부구분 · Media · Campaign · Ad Group · 종료일\n\n"
                   "업로드 후에는 다시 올리지 않아도 계속 적용됩니다.")

    st.divider()

    # ── 개별 가공 버튼
    b1, b2 = st.columns(2)
    with b1:
        run_media = st.button("📊 미디어 파일 가공", type="primary", use_container_width=True)
    with b2:
        run_cat = st.button("🛒 카테고리 파일 가공", type="primary", use_container_width=True)

    group_df = st.session_state['index_group_df']
    lookup   = build_index_lookup(group_df)

    # ────────────────────────────── 미디어 가공 ──────────────────────────────
    if run_media:
        if not media_file:
            st.warning("미디어 파일을 업로드해주세요.")
        else:
            try:
                media_raw = read_csv_robust(media_file)
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")
                media_raw = None

            if media_raw is not None:
                if detect_file_kind(media_raw) == 'category':
                    st.warning(f"⚠️ '{media_file.name}'은(는) 카테고리 파일 컬럼 구조로 보입니다. "
                               "카테고리 파일 업로드란에 올린 뒤 카테고리 가공 버튼을 사용해주세요.")
                else:
                    if media_raw.shape[1] != len(MEDIA_COLUMNS):
                        st.info(f"ℹ️ 미디어 파일 컬럼 수({media_raw.shape[1]})가 예상"
                                f"({len(MEDIA_COLUMNS)}개)과 다릅니다. 컬럼 구조를 확인해주세요.")
                    with st.spinner("미디어 데이터 가공 중..."):
                        try:
                            m_df, m_d47 = process_media(media_raw, yesterday)
                            if m_df is not None:
                                date_col = m_df.columns[M_DATE]
                                camp_col = m_df.columns[M_CAMP]
                                adg_col  = m_df.columns[M_ADG]
                                if lookup:
                                    m_df = add_index_columns(m_df, lookup, camp_col, adg_col, date_col)
                                else:
                                    st.info("ℹ️ 적용된 인덱스가 없어 사업부구분(인덱스)/D7_상태 열은 "
                                            "추가되지 않았습니다.")
                                st.session_state['media_df']   = m_df
                                st.session_state['media_d4_7'] = m_d47
                        except Exception:
                            st.error("미디어 처리 오류:\n```\n" + traceback.format_exc() + "\n```")

    # ────────────────────────────── 카테고리 가공 ────────────────────────────
    if run_cat:
        if not cat_file:
            st.warning("카테고리 파일을 업로드해주세요.")
        else:
            try:
                cat_raw = read_csv_robust(cat_file)
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")
                cat_raw = None

            if cat_raw is not None:
                if detect_file_kind(cat_raw) == 'media':
                    st.warning(f"⚠️ '{cat_file.name}'은(는) 미디어 파일 컬럼 구조로 보입니다. "
                               "미디어 파일 업로드란에 올린 뒤 미디어 가공 버튼을 사용해주세요.")
                else:
                    if cat_raw.shape[1] != len(CATEGORY_COLUMNS):
                        st.info(f"ℹ️ 카테고리 파일 컬럼 수({cat_raw.shape[1]})가 예상"
                                f"({len(CATEGORY_COLUMNS)}개)과 다릅니다. 컬럼 구조를 확인해주세요.")
                    if not lookup:
                        st.warning("⚠️ 적용된 인덱스가 없어 사업부구분을 판별할 수 없습니다. "
                                   "카테고리_unique/qty/price가 모두 '카테고리 구매' 총계로 대체되고 "
                                   "전체 행이 수기 확인 대상으로 표시됩니다. 인덱스 파일을 먼저 업로드해주세요.")
                    with st.spinner("카테고리 데이터 가공 중..."):
                        try:
                            c_df, c_manual = process_category(cat_raw, yesterday, lookup)
                            st.session_state['cat_df']     = c_df
                            st.session_state['cat_manual'] = c_manual
                        except Exception:
                            st.error("카테고리 처리 오류:\n```\n" + traceback.format_exc() + "\n```")

    if not run_media and not run_cat and st.session_state['media_df'] is None \
            and st.session_state['cat_df'] is None:
        st.info("파일을 업로드한 후 원하는 가공 버튼을 눌러주세요.")

    media_df   = st.session_state['media_df']
    media_d4_7 = st.session_state['media_d4_7']
    cat_df     = st.session_state['cat_df']
    cat_manual = st.session_state['cat_manual']

    st.divider()

    # ────────────────────────────── 미디어 결과 ──────────────────────────────
    if media_df is not None and not media_df.empty:
        st.subheader("📊 미디어 가공 결과")

        n_d4_7 = int(media_d4_7.sum()) if media_d4_7 is not None else 0
        m1, m2, m3 = st.columns(3)
        m1.metric("전체 행",            len(media_df))
        m2.metric("1~3일치 (전체 열)",  len(media_df) - n_d4_7)
        m3.metric("4~7일치 (BL열만 유효)", n_d4_7)

        # 미리보기: 인덱스 열 + 식별자 열(Date~ad) + 집약형(Adef)열
        id_cols = [c for c in ['사업부구분(인덱스)', 'D7_상태'] if c in media_df.columns]
        id_cols += [c for c in MEDIA_COLUMNS[:8] if c in media_df.columns]
        if M_BL_NAME in media_df.columns and M_BL_NAME not in id_cols:
            id_cols.append(M_BL_NAME)

        preview   = media_df[id_cols].copy()
        mask_prev = media_d4_7[:len(preview)] if media_d4_7 is not None else None

        st.dataframe(
            highlight_rows(preview, mask_prev, '#FFF3CD'),
            use_container_width=True, height=320
        )
        st.caption("🟡 노란 행 = 4~7일치: I열~BK열 수치 0 처리됨, BL열 원본 유지  "
                   "| 전체 열은 다운로드 파일에 포함됩니다.")

    # ────────────────────────────── 카테고리 결과 ────────────────────────────
    if cat_df is not None and not cat_df.empty:
        st.subheader("🛒 카테고리 가공 결과")

        n_m = int(cat_manual.sum()) if cat_manual is not None else 0
        if n_m:
            st.warning(f"⚠️ 수기 확인 필요: **{n_m}개 행** — "
                       "사업부구분 미매핑 또는 사업부-연합의 USP Category 미매핑")
        else:
            st.success("모든 행이 정상 매핑되었습니다.")

        st.dataframe(
            highlight_rows(cat_df, cat_manual, '#F8D7DA'),
            use_container_width=True, height=320
        )
        st.caption("🔴 빨간 행 = 수기 확인 필요  "
                   "| 다운로드 시 카테고리_RD붙여넣기 시트가 RD 시트 BU~BW 붙여넣기용입니다.")

        # 수기 확인 필요 행 별도 표시
        if n_m:
            with st.expander(f"🔍 수기 확인 필요 항목 ({n_m}개) 자세히 보기"):
                manual_rows = cat_df[cat_manual.astype(bool)].copy() if cat_manual is not None else pd.DataFrame()
                st.dataframe(manual_rows, use_container_width=True)

    # ────────────────────────────── 다운로드 ─────────────────────────────────
    st.divider()
    has_output = (media_df is not None and not media_df.empty) or \
                 (cat_df   is not None and not cat_df.empty)

    if has_output:
        st.subheader("💾 가공 데이터 다운로드")
        fname       = f"NPS_가공_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
        excel_bytes = build_excel(media_df, cat_df)

        st.download_button(
            "📥 엑셀 다운로드",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        sheets = []
        if media_df is not None and not media_df.empty:
            sheets.append("**미디어_가공**: 7일치 전체 열 (4~7일치 수치 0 처리)")
        if cat_df is not None and not cat_df.empty:
            sheets.append("**카테고리_RD붙여넣기**: RD 시트 BU~BW 붙여넣기용")
            sheets.append("**카테고리_전체**: 수기확인 플래그 포함 전체 데이터")

        st.info("시트 구성  |  " + "  |  ".join(sheets))
    else:
        st.info("가공된 데이터가 없습니다. 파일을 확인 후 다시 시도해주세요.")


if __name__ == "__main__":
    main()

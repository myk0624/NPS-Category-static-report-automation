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

# Category file
C_DATE, C_MEDIA, C_CAMP, C_ADG = ci('A'), ci('E'), ci('F'), ci('G')
C_AV = ci('AV')                  # 47 – USP Category (사업부-연합 분기)
C_BR = ci('BR')                  # 69 – 사업부구분
C_BS, C_BT, C_BU = ci('BS'), ci('BT'), ci('BU')  # 70, 71, 72

# Replacement-source columns per 사업부: (unique_col, qty_col, price_col)
BIZ_MAP = {
    '사업부-가구':          (ci('CE'), ci('CF'), ci('CG')),
    '사업부-그로서리-전체': (ci('CK'), ci('CL'), ci('CM')),
    '사업부-그로서리-별도': (ci('CK'), ci('CL'), ci('CM')),
    '사업부-리빙':          (ci('CW'), ci('CX'), ci('CY')),
    '사업부-자동차공구':    (ci('CZ'), ci('DA'), ci('DB')),
    '사업부-키즈':          (ci('CH'), ci('CI'), ci('CJ')),
    '사업부-펫':            (ci('DC'), ci('DD'), ci('DE')),
    '사업부-여가생활e쿠폰': (ci('DI'), ci('DJ'), ci('DK')),
}
USP_MAP = {
    'LVG':     (ci('CW'), ci('CX'), ci('CY')),
    'PET':     (ci('DC'), ci('DD'), ci('DE')),
    'CARTOOL': (ci('CZ'), ci('DA'), ci('DB')),
    'KID':     (ci('CH'), ci('CI'), ci('CJ')),
}


# ──────────────────────────────────────────────────────────────────────────────
# Utilities
# ──────────────────────────────────────────────────────────────────────────────

def safe_col(df, idx):
    return df.columns[idx] if idx < len(df.columns) else None


def row_val(row, idx, default=np.nan):
    return row.iloc[idx] if idx < len(row) else default


# ──────────────────────────────────────────────────────────────────────────────
# D7 end-date filter
# ──────────────────────────────────────────────────────────────────────────────

def apply_d7_filter(data, end_df, camp_col, adg_col, date_col):
    """Keep rows where data_date ≤ campaign_end_date + 7 days.
    Rows not found in end_df are kept as-is."""
    if end_df is None or end_df.empty:
        return data
    ec = end_df.columns.tolist()
    if len(ec) < 2:
        return data

    has_adg      = len(ec) >= 3
    end_date_col = ec[2] if has_adg else ec[1]

    end = end_df.copy()
    end[end_date_col] = pd.to_datetime(end[end_date_col], errors='coerce')
    end = end.dropna(subset=[end_date_col])

    if has_adg:
        lookup = {
            (str(r[ec[0]]).strip(), str(r[ec[1]]).strip()): r[end_date_col]
            for _, r in end.iterrows()
        }
    else:
        lookup = {str(r[ec[0]]).strip(): r[end_date_col] for _, r in end.iterrows()}

    def keep(row):
        k = (str(row[camp_col]).strip(), str(row[adg_col]).strip()) if has_adg \
            else str(row[camp_col]).strip()
        ed = lookup.get(k)
        return ed is None or row[date_col] <= ed + pd.Timedelta(days=7)

    return data[data.apply(keep, axis=1)].reset_index(drop=True)


# ──────────────────────────────────────────────────────────────────────────────
# Media processing
# ──────────────────────────────────────────────────────────────────────────────

def process_media(raw, end_df, yesterday):
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
    camp_col = df.columns[M_CAMP]
    adg_col  = df.columns[M_ADG]

    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col]).sort_values(date_col).reset_index(drop=True)

    d7_start = yesterday - pd.Timedelta(days=6)
    d3_start = yesterday - pd.Timedelta(days=2)  # 3일치 중 가장 오래된 날

    df = df[(df[date_col] >= d7_start) & (df[date_col] <= yesterday)].copy()
    df = df.reset_index(drop=True)

    if end_df is not None and not end_df.empty:
        df = apply_d7_filter(df, end_df, camp_col, adg_col, date_col)

    is_d4_7 = (df[date_col] < d3_start).values  # bool array aligned with df

    # Zero out I(8)..BK(62) for day-4-to-7 rows; BL(63) kept
    for col_i in range(M_NUM_ST, M_BL):
        if col_i < df.shape[1]:
            df.iloc[is_d4_7, col_i] = 0

    return df, is_d4_7


# ──────────────────────────────────────────────────────────────────────────────
# Category processing
# ──────────────────────────────────────────────────────────────────────────────

def process_category(raw, end_df, yesterday):
    """
    Returns (result_df, needs_manual_bool_array).
    result_df columns: 날짜, Media, Campaign, Ad Group, 사업부구분, USP_Category,
                       카테고리_unique, 카테고리_qty, 카테고리_price, 수기확인필요
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
    if end_df is not None and not end_df.empty and camp_col and adg_col:
        df = apply_d7_filter(df, end_df, camp_col, adg_col, date_col)

    n      = len(df)
    out_u  = np.full(n, np.nan, object)
    out_q  = np.full(n, np.nan, object)
    out_p  = np.full(n, np.nan, object)
    manual = np.zeros(n, bool)

    for pos, (_, row) in enumerate(df.iterrows()):
        biz = str(row_val(row, C_BR, '')).strip()

        if biz == '사업부-연합':
            usp  = str(row_val(row, C_AV, '')).strip().upper()
            cols = USP_MAP.get(usp)
            if cols:
                out_u[pos] = row_val(row, cols[0])
                out_q[pos] = row_val(row, cols[1])
                out_p[pos] = row_val(row, cols[2])
            else:
                # USP 값 미매핑 → 원본 유지 + 수기 확인 플래그
                out_u[pos]  = row_val(row, C_BS)
                out_q[pos]  = row_val(row, C_BT)
                out_p[pos]  = row_val(row, C_BU)
                manual[pos] = True

        elif biz in BIZ_MAP:
            c_u, c_q, c_p = BIZ_MAP[biz]
            out_u[pos] = row_val(row, c_u)
            out_q[pos] = row_val(row, c_q)
            out_p[pos] = row_val(row, c_p)

        else:
            # 알 수 없는 사업부구분 → 원본 유지 + 수기 확인 플래그
            out_u[pos]  = row_val(row, C_BS)
            out_q[pos]  = row_val(row, C_BT)
            out_p[pos]  = row_val(row, C_BU)
            if biz and biz.lower() not in ('nan', ''):
                manual[pos] = True

    result = pd.DataFrame({
        '날짜':             df.iloc[:, C_DATE].dt.date      if C_DATE  < df.shape[1] else '',
        'Media':            df.iloc[:, C_MEDIA]             if C_MEDIA < df.shape[1] else '',
        'Campaign':         df.iloc[:, C_CAMP]              if C_CAMP  < df.shape[1] else '',
        'Ad Group':         df.iloc[:, C_ADG]               if C_ADG   < df.shape[1] else '',
        '사업부구분':       df.iloc[:, C_BR]                if C_BR    < df.shape[1] else '',
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

def build_excel(media_df, cat_df, end_df):
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

        if end_df is not None and not end_df.empty:
            end_df.to_excel(writer, sheet_name='캠페인종료일', index=False)

    buf.seek(0)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="NPS Report 가공기", layout="wide", page_icon="📊")
    st.title("📊 NPS Report 데이터 가공기")
    st.caption("미디어·카테고리 CSV 로우 데이터를 업로드하면 RD 시트 형식으로 자동 가공합니다.")

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
    st.caption("파일명과 무관하게 컬럼 구조를 읽어 미디어/카테고리 파일을 자동으로 판별·교정합니다.")

    with c3:
        st.markdown("### 📅 캠페인 종료일 (선택)")
        end_file = st.file_uploader("종료일 xlsx 업로드", type=['xlsx', 'xls'], key='ef')
        if end_file:
            st.success(f"✅ {end_file.name}")
        st.caption("열 구조: **A** Campaign | **B** Ad Group | **C** End Date\n\n"
                   "종료일 + 7일 이후 데이터는 자동 제외됩니다.")

    st.divider()
    run = st.button("🔄 가공 시작", type="primary", use_container_width=True)

    if not run:
        st.info("파일을 업로드한 후 **가공 시작** 버튼을 눌러주세요.")
        return

    if not media_file and not cat_file:
        st.warning("미디어 파일 또는 카테고리 파일을 업로드해주세요.")
        return

    # ── 파일 읽기
    with st.spinner("파일 읽는 중..."):
        try:
            media_raw = read_csv_robust(media_file) if media_file else None
            cat_raw   = read_csv_robust(cat_file)   if cat_file   else None
            end_df    = pd.read_excel(end_file, header=0, engine='openpyxl') \
                if end_file   else None
        except Exception as e:
            st.error(f"파일 읽기 오류: {e}")
            return

    # ── 컬럼 구조 기반 파일 종류 자동 판별/교정
    # (파일명이 아닌 '카테고리 구매 unique' 컬럼 존재 여부로 판단)
    media_kind = detect_file_kind(media_raw)
    cat_kind   = detect_file_kind(cat_raw)

    if media_kind == 'category' and cat_kind == 'media':
        st.warning("⚠️ 업로드 위치가 바뀐 것을 감지하여 자동으로 교정했습니다: "
                   f"'{media_file.name}' → 카테고리 파일, '{cat_file.name}' → 미디어 파일")
        media_raw, cat_raw = cat_raw, media_raw
    elif media_kind == 'category' and cat_raw is None:
        st.warning(f"⚠️ '{media_file.name}'은(는) 카테고리 파일 컬럼 구조로 감지되어 "
                   "카테고리 파일로 처리합니다.")
        cat_raw, media_raw = media_raw, None
    elif cat_kind == 'media' and media_raw is None:
        st.warning(f"⚠️ '{cat_file.name}'은(는) 미디어 파일 컬럼 구조로 감지되어 "
                   "미디어 파일로 처리합니다.")
        media_raw, cat_raw = cat_raw, None

    if media_raw is not None and media_raw.shape[1] not in (len(MEDIA_COLUMNS),):
        st.info(f"ℹ️ 미디어 파일 컬럼 수({media_raw.shape[1]})가 예상({len(MEDIA_COLUMNS)}개)과 "
                "다릅니다. 컬럼 구조를 확인해주세요.")
    if cat_raw is not None and cat_raw.shape[1] not in (len(CATEGORY_COLUMNS),):
        st.info(f"ℹ️ 카테고리 파일 컬럼 수({cat_raw.shape[1]})가 예상({len(CATEGORY_COLUMNS)}개)과 "
                "다릅니다. 컬럼 구조를 확인해주세요.")

    # ── 가공
    media_df = media_d4_7 = cat_df = cat_manual = None

    if media_raw is not None:
        with st.spinner("미디어 데이터 가공 중..."):
            try:
                media_df, media_d4_7 = process_media(media_raw, end_df, yesterday)
            except Exception:
                st.error("미디어 처리 오류:\n```\n" + traceback.format_exc() + "\n```")

    if cat_raw is not None:
        with st.spinner("카테고리 데이터 가공 중..."):
            try:
                cat_df, cat_manual = process_category(cat_raw, end_df, yesterday)
            except Exception:
                st.error("카테고리 처리 오류:\n```\n" + traceback.format_exc() + "\n```")

    st.divider()

    # ────────────────────────────── 미디어 결과 ──────────────────────────────
    if media_df is not None and not media_df.empty:
        st.subheader("📊 미디어 가공 결과")

        n_d4_7 = int(media_d4_7.sum()) if media_d4_7 is not None else 0
        m1, m2, m3 = st.columns(3)
        m1.metric("전체 행",            len(media_df))
        m2.metric("1~3일치 (전체 열)",  len(media_df) - n_d4_7)
        m3.metric("4~7일치 (BL열만 유효)", n_d4_7)

        # 미리보기: 식별자 열(A~H) + BL열
        show = list(media_df.columns[:min(8, len(media_df.columns))])
        if M_BL < len(media_df.columns) and media_df.columns[M_BL] not in show:
            show.append(media_df.columns[M_BL])

        preview   = media_df[show].copy()
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

    # ────────────────────────────── 캠페인 종료일 ────────────────────────────
    if end_df is not None and not end_df.empty:
        st.subheader("📅 캠페인 종료일")

        ec       = end_df.columns.tolist()
        date_ec  = ec[2] if len(ec) >= 3 else ec[-1]
        ed       = end_df.copy()
        ed[date_ec] = pd.to_datetime(ed[date_ec], errors='coerce')

        delta = (ed[date_ec] - yesterday).dt.days
        ended  = (delta < 0).values
        soon   = ((delta >= 0) & (delta <= 3)).values

        color_arr = np.where(ended, '#F8D7DA', np.where(soon, '#FFF3CD', ''))

        def _end_style(x):
            bg = pd.DataFrame('', index=x.index, columns=x.columns)
            for i, color in enumerate(color_arr):
                if color and i < len(x):
                    bg.iloc[i] = f'background-color: {color}'
            return bg

        st.dataframe(ed.style.apply(_end_style, axis=None), use_container_width=True)
        st.caption(f"🔴 이미 종료 ({int(ended.sum())}개)  |  "
                   f"🟡 3일 이내 종료 예정 ({int(soon.sum())}개)  |  "
                   f"⬜ 정상 운영 중")

    # ────────────────────────────── 다운로드 ─────────────────────────────────
    st.divider()
    has_output = (media_df is not None and not media_df.empty) or \
                 (cat_df   is not None and not cat_df.empty)

    if has_output:
        st.subheader("💾 가공 데이터 다운로드")
        fname       = f"NPS_가공_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
        excel_bytes = build_excel(media_df, cat_df, end_df)

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
        if end_df is not None and not end_df.empty:
            sheets.append("**캠페인종료일**: 업로드된 종료일 원본")

        st.info("시트 구성  |  " + "  |  ".join(sheets))
    else:
        st.info("가공된 데이터가 없습니다. 파일을 확인 후 다시 시도해주세요.")


if __name__ == "__main__":
    main()

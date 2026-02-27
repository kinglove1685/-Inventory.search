import streamlit as st
from pathlib import Path
from datetime import datetime, timezone, timedelta
from io import BytesIO
import json
import os
import re
import pandas as pd

from inventory_search import load_inventory, load_inventory_from_bytes, search_inventory, summarize_inventory, lot_breakdown
from inventory_search import (
    build_summary_export_multi,
    COL_COLOR,
    COL_COLOR_HEX,
    COL_LOTNO,
    COL_ITEM,
    COL_NAME,
    COL_P,
    COL_T,
    COL_POWER,
    COL_TONE,
    COL_CYL,
    COL_AXIS,
    COL_ADD,
)

DEFAULT_PATH = Path(__file__).parent / "재고관련 프로그램제작.xlsx"
LIGHTWEIGHT_DEFAULT_PATH = Path(__file__).parent / "재고장_더미.xlsx"
LIGHTWEIGHT_ASCII_PATH = Path(__file__).parent / "inventory_dummy.xlsx"

try:
    from streamlit_paste_button import paste_image_button as _paste_image_button
except Exception:
    _paste_image_button = None

TEST_MODE = os.environ.get("INVENTORY_APP_TEST", "").strip().lower() in {"1", "true", "yes", "y"}
default_page_title = "재고 검색 TEST" if TEST_MODE else "재고 검색"
default_dashboard_title = "재고 검색 대시보드"
PAGE_TITLE = os.environ.get("INVENTORY_APP_PAGE_TITLE", default_page_title)
DASHBOARD_TITLE = os.environ.get("INVENTORY_APP_DASHBOARD_TITLE", default_dashboard_title)
st.set_page_config(page_title=PAGE_TITLE, layout="wide")

st.markdown(
    """
    <style>
    :root {
        --app-font: "Pretendard", "SUIT", "Noto Sans KR", "Segoe UI", sans-serif;
        --text-strong: #1e2a44;
        --text-base: #2f3b55;
        --text-soft: #6b7895;
    }
    html, body, [data-testid="stAppViewContainer"], section[data-testid="stSidebar"] {
        font-family: var(--app-font);
    }
    [data-testid="stAppViewContainer"] .main .block-container {
        padding-top: 1.2rem;
    }
    .material-icons,
    .material-icons-outlined,
    .material-icons-round,
    .material-icons-sharp,
    .material-symbols-outlined,
    .material-symbols-rounded,
    .material-symbols-sharp {
        font-family: "Material Icons", "Material Symbols Outlined", "Material Symbols Rounded", "Material Symbols Sharp" !important;
        font-weight: normal !important;
        font-style: normal !important;
        letter-spacing: normal !important;
        text-transform: none !important;
        white-space: nowrap !important;
        direction: ltr !important;
    }
    [data-testid="stAppViewContainer"] h1 {
        font-size: 2.56rem;
        line-height: 1.15;
        font-weight: 800;
        letter-spacing: -0.01em;
        color: var(--text-strong);
    }
    [data-testid="stAppViewContainer"] h2 {
        font-size: 1.85rem;
        font-weight: 750;
        letter-spacing: -0.005em;
        color: var(--text-strong);
    }
    [data-testid="stAppViewContainer"] h3 {
        font-size: 1.56rem;
        font-weight: 700;
        color: var(--text-strong);
    }
    [data-testid="stCheckbox"] label p {
        font-size: 1.28rem;
        font-weight: 650;
    }
    [data-testid="stAppViewContainer"] label,
    [data-testid="stAppViewContainer"] .stMarkdown,
    [data-testid="stAppViewContainer"] p {
        color: var(--text-base);
        font-size: 0.97rem;
    }
    [data-testid="stAppViewContainer"] [data-testid="stCaptionContainer"] {
        color: var(--text-soft);
        font-size: 0.83rem;
    }
    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input,
    [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        font-size: 0.98rem;
        font-weight: 550;
    }
    [data-testid="stAppViewContainer"] [data-testid="stDataFrame"] {
        font-size: 0.94rem;
    }
    section[data-testid="stSidebar"] > div {
        background:
            radial-gradient(120% 100% at 0% 0%, rgba(61, 121, 242, 0.18) 0%, rgba(61, 121, 242, 0.00) 52%),
            linear-gradient(180deg, #edf2ff 0%, #e3ebff 100%);
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1.1rem;
        padding-bottom: 1.1rem;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"] {
        border: 1px solid #d4def4;
        border-radius: 16px;
        background: rgba(255, 255, 255, 0.78);
        box-shadow: 0 10px 24px rgba(30, 64, 175, 0.08);
        margin-bottom: 0.85rem;
        overflow: hidden;
        backdrop-filter: blur(6px);
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"] summary {
        background: rgba(255, 255, 255, 0.96);
        border-bottom: 1px solid #e4ebfb;
        border-radius: 16px 16px 0 0;
        padding: 0.55rem 0.72rem;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"] summary p {
        font-size: 1.12rem;
        font-weight: 800;
        color: #1e2a44;
        letter-spacing: 0.01em;
    }
    section[data-testid="stSidebar"] .stButton > button {
        border-radius: 11px;
        width: 100%;
        text-align: left;
        padding: 0.56rem 0.76rem;
        border: 1px solid transparent;
        background: rgba(255, 255, 255, 0.6);
        transition: all 120ms ease;
        font-size: 1rem;
    }
    section[data-testid="stSidebar"] .stButton > button:hover {
        border-color: #c7d6fb;
        background: rgba(255, 255, 255, 0.98);
    }
    section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2f7cf3 0%, #4c8ff7 100%);
        border-color: #2f7cf3;
        color: #ffffff !important;
        font-weight: 800;
        box-shadow: 0 8px 18px rgba(47, 124, 243, 0.32);
    }
    section[data-testid="stSidebar"] .stButton > button[kind="primary"] * {
        color: #ffffff !important;
        fill: #ffffff !important;
    }
    section[data-testid="stSidebar"] .stButton > button[kind="secondary"],
    section[data-testid="stSidebar"] .stButton > button[kind="tertiary"] {
        font-weight: 700;
        color: #32405f;
    }
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
        color: #64708e;
        font-size: 0.86rem;
    }
    section[data-testid="stSidebar"] .stButton {
        margin-bottom: 0.3rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

title_col, meta_col = st.columns([3, 1])
with title_col:
    st.title(DASHBOARD_TITLE)

with st.sidebar:
    st.header("메뉴")
    if "active_screen" not in st.session_state:
        st.session_state["active_screen"] = "integrated"

    with st.expander("검색 표시", expanded=True):
        if st.button(
            "통합 검색",
            key="nav_integrated",
            use_container_width=True,
            type="primary" if st.session_state["active_screen"] == "integrated" else "tertiary",
        ):
            st.session_state["active_screen"] = "integrated"
            st.rerun()
        if st.button(
            "OCR 검색",
            key="nav_ocr",
            use_container_width=True,
            type="primary" if st.session_state["active_screen"] == "ocr" else "tertiary",
        ):
            st.session_state["active_screen"] = "ocr"
            st.rerun()

    with st.expander("멸균넘버 기준표", expanded=True):
        if st.button(
            "멸균넘버기준표",
            key="nav_sterile_ref",
            use_container_width=True,
            type="primary" if st.session_state["active_screen"] == "sterile_ref" else "tertiary",
        ):
            st.session_state["active_screen"] = "sterile_ref"
            st.rerun()

active_screen = st.session_state.get("active_screen", "integrated")
show_search = active_screen == "integrated"
show_ocr = active_screen == "ocr"
show_sterile_ref = active_screen == "sterile_ref"

_persist_defaults = {
    "toric_mf": False,
    "export_no_power": False,
    "query": "",
    "color_query": "",
    "tone_query": "",
    "power_query": "",
    "cyl_query": "",
    "axis_query": "",
    "add_query": "",
    "category_query": "전체",
    "sterile_query": "",
    "sterile_mode": "고정",
    "ocr_query": "",
    "ocr_text": "",
}
for _k, _v in _persist_defaults.items():
    _pk = f"persist_{_k}"
    if _pk not in st.session_state:
        st.session_state[_pk] = _v
    if _k in st.session_state:
        st.session_state[_pk] = st.session_state[_k]
    else:
        st.session_state[_k] = st.session_state[_pk]

@st.cache_data(show_spinner=False)
def load_from_path(path: Path) -> pd.DataFrame:
    return load_inventory(path, load_color=False)

@st.cache_data(show_spinner=False)
def load_from_bytes(data: bytes) -> pd.DataFrame:
    return load_inventory_from_bytes(data, load_color=False)


@st.cache_data(show_spinner=False)
def load_sterile_reference(path: Path) -> pd.DataFrame:
    empty_df = pd.DataFrame(columns=["시작일자", "종료일자", "멸균LOT", "유효기간(5년)", "유효기간(8년)"])
    try:
        xls = pd.ExcelFile(path)
    except Exception:
        return empty_df

    sheet_name = None
    preferred_names = ["멸균넘버 기준표", "멸균넘버기준표", "멸균 기준표", "멸균기준표"]
    for name in preferred_names:
        if name in xls.sheet_names:
            sheet_name = name
            break
    if sheet_name is None:
        for name in xls.sheet_names:
            n = str(name).replace(" ", "")
            if "멸균" in n and "기준" in n:
                sheet_name = name
                break
    if sheet_name is None:
        return empty_df

    try:
        ref_df = pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return empty_df
    lot_col = None
    for c in ref_df.columns:
        name = str(c).strip()
        if name in {"멸균LOT", "멸균넘버", "멸균No.", "멸균NO.", "LOTNO"}:
            lot_col = c
            break
    if lot_col is None:
        return empty_df
    ref_df = ref_df.copy()
    ref_df = ref_df.rename(columns={lot_col: "멸균LOT"})
    if "멸균LOT" in ref_df.columns:
        ref_df["멸균LOT"] = ref_df["멸균LOT"].fillna("").astype(str).str.strip().str.upper()
        ref_df = ref_df[ref_df["멸균LOT"] != ""]
    keep_cols = ["시작일자", "종료일자", "멸균LOT", "유효기간(5년)", "유효기간(8년)"]
    for c in keep_cols:
        if c not in ref_df.columns:
            ref_df[c] = ""
    return ref_df[keep_cols]


def _sterile_code_to_rank(code: str) -> int | None:
    s = str(code or "").strip().upper()
    m = re.fullmatch(r"([A-Z]{2})(\d{2})", s)
    if not m:
        return None
    a, b = m.group(1)[0], m.group(1)[1]
    prefix_rank = (ord(a) - 65) * 26 + (ord(b) - 65)
    return prefix_rank * 100 + int(m.group(2))


def _sterile_rank_to_code(rank: int) -> str | None:
    if rank < 0:
        return None
    prefix_rank, num = divmod(rank, 100)
    if prefix_rank >= 26 * 26:
        return None
    a = chr(65 + (prefix_rank // 26))
    b = chr(65 + (prefix_rank % 26))
    return f"{a}{b}{num:02d}"


def _build_sterile_order_map(ref_df: pd.DataFrame) -> dict[str, int]:
    order_map: dict[str, int] = {}
    for v in ref_df.get("멸균LOT", pd.Series(dtype=str)):
        code = str(v).strip().upper()
        if not code:
            continue
        rank = _sterile_code_to_rank(code)
        if rank is not None:
            order_map[code] = rank
        elif code not in order_map:
            order_map[code] = len(order_map)
    return order_map


def _build_sterile_reference_view(ref_df: pd.DataFrame, before_count: int = 20, after_count: int = 20) -> pd.DataFrame:
    view_df = ref_df.copy()
    if view_df.empty:
        return view_df
    # Normalize date columns for calculations (keeps original columns for display).
    for col in ["시작일자", "종료일자", "유효기간(5년)", "유효기간(8년)"]:
        if col in view_df.columns:
            view_df[col] = pd.to_datetime(view_df[col], errors="coerce")

    ranks = []
    for code in view_df["멸균LOT"].astype(str):
        rank = _sterile_code_to_rank(code)
        if rank is not None:
            ranks.append(rank)
    if not ranks:
        return view_df

    ranked_df = view_df.copy()
    ranked_df["_rank"] = ranked_df["멸균LOT"].astype(str).map(_sterile_code_to_rank)
    ranked_df = ranked_df[ranked_df["_rank"].notna()].sort_values("_rank")
    dated_df = ranked_df.dropna(subset=["시작일자", "종료일자"]).copy()

    # Determine the most common step (start date delta) and span (end-start).
    step_days = 14
    span_days = 13
    if not dated_df.empty:
        start_delta = dated_df["시작일자"].diff().dt.days
        span_delta = (dated_df["종료일자"] - dated_df["시작일자"]).dt.days
        if not start_delta.dropna().empty:
            step_days = int(start_delta.dropna().value_counts().idxmax())
        if not span_delta.dropna().empty:
            span_days = int(span_delta.dropna().value_counts().idxmax())

    def _calc_expiry(start_date: pd.Timestamp, years: int) -> pd.Timestamp | str:
        if pd.isna(start_date):
            return ""
        base = start_date + pd.DateOffset(years=years, months=-1)
        return (base + pd.offsets.MonthEnd(0)).normalize()

    min_rank = min(ranks)
    max_rank = max(ranks)
    prev_rows = []
    next_rows = []

    base_min_start = None
    base_max_start = None
    if not dated_df.empty:
        base_min_start = dated_df.iloc[0]["시작일자"]
        base_max_start = dated_df.iloc[-1]["시작일자"]

    for rank in range(max(0, min_rank - before_count), min_rank):
        code = _sterile_rank_to_code(rank)
        if code:
            row = {"멸균LOT": code}
            if base_min_start is not None:
                delta = rank - min_rank
                start = base_min_start + pd.Timedelta(days=step_days * delta)
                end = start + pd.Timedelta(days=span_days)
                row["시작일자"] = start
                row["종료일자"] = end
                row["유효기간(5년)"] = _calc_expiry(start, 5)
                row["유효기간(8년)"] = _calc_expiry(start, 8)
            prev_rows.append(row)
    for rank in range(max_rank + 1, max_rank + after_count + 1):
        code = _sterile_rank_to_code(rank)
        if code:
            row = {"멸균LOT": code}
            if base_max_start is not None:
                delta = rank - max_rank
                start = base_max_start + pd.Timedelta(days=step_days * delta)
                end = start + pd.Timedelta(days=span_days)
                row["시작일자"] = start
                row["종료일자"] = end
                row["유효기간(5년)"] = _calc_expiry(start, 5)
                row["유효기간(8년)"] = _calc_expiry(start, 8)
            next_rows.append(row)

    prev_df = pd.DataFrame(prev_rows)
    next_df = pd.DataFrame(next_rows)
    for col in ["시작일자", "종료일자", "유효기간(5년)", "유효기간(8년)"]:
        if col not in prev_df.columns:
            prev_df[col] = ""
        if col not in next_df.columns:
            next_df[col] = ""
    ordered_cols = ["시작일자", "종료일자", "멸균LOT", "유효기간(5년)", "유효기간(8년)"]
    result_df = pd.concat(
        [
            prev_df[ordered_cols],
            view_df[ordered_cols],
            next_df[ordered_cols],
        ],
        ignore_index=True,
    )
    for col in ["시작일자", "종료일자", "유효기간(5년)", "유효기간(8년)"]:
        if col in result_df.columns:
            result_df[col] = pd.to_datetime(result_df[col], errors="coerce").dt.strftime("%Y-%m-%d")
    return result_df


def _get_service_account_info() -> dict | None:
    if "gcp_service_account" in st.secrets:
        try:
            return dict(st.secrets["gcp_service_account"])
        except Exception:
            return st.secrets["gcp_service_account"]
    if "gcp_service_account_json" in st.secrets:
        try:
            return json.loads(st.secrets["gcp_service_account_json"])
        except Exception:
            return None
    return None


def _get_vision_client():
    from google.cloud import vision
    from google.oauth2 import service_account

    info = _get_service_account_info()
    if not info:
        return None
    creds = service_account.Credentials.from_service_account_info(info)
    return vision.ImageAnnotatorClient(credentials=creds)


def _image_to_bytes(image_obj) -> bytes | None:
    if image_obj is None:
        return None
    if isinstance(image_obj, (bytes, bytearray)):
        return bytes(image_obj)
    if hasattr(image_obj, "getvalue"):
        try:
            return image_obj.getvalue()
        except Exception:
            pass
    try:
        from PIL import Image

        if isinstance(image_obj, Image.Image):
            buffer = BytesIO()
            image_obj.save(buffer, format="PNG")
            return buffer.getvalue()
    except Exception:
        pass
    return None


def run_ocr(image_bytes: bytes) -> str:
    client = _get_vision_client()
    if client is None:
        raise RuntimeError("GCP 서비스 계정 정보가 설정되지 않았습니다.")
    from google.cloud import vision

    image = vision.Image(content=image_bytes)
    response = client.document_text_detection(
        image=image,
        image_context={"language_hints": ["ko", "en"]},
    )
    if response.error.message:
        raise RuntimeError(response.error.message)
    if response.full_text_annotation and response.full_text_annotation.text:
        return response.full_text_annotation.text
    return ""

# Load data
source_bytes = None
source_path = None
with st.sidebar:
    inventory_upload = st.file_uploader(
        "재고 엑셀 업로드 (.xlsx)",
        type=["xlsx", "xlsm"],
        help="업로드하면 해당 파일의 재고장 데이터를 사용합니다. 업로드하지 않으면 기본 파일을 사용합니다.",
    )

if DEFAULT_PATH.exists():
    default_source = DEFAULT_PATH
elif LIGHTWEIGHT_DEFAULT_PATH.exists():
    default_source = LIGHTWEIGHT_DEFAULT_PATH
else:
    default_source = LIGHTWEIGHT_ASCII_PATH
if inventory_upload is not None:
    source_bytes = inventory_upload.getvalue()
    df = load_from_bytes(source_bytes)
    sterile_ref_df = pd.DataFrame(columns=["시작일자", "종료일자", "멸균LOT", "유효기간(5년)", "유효기간(8년)"])
elif default_source.exists():
    df = load_from_path(default_source)
    source_path = default_source
    sterile_ref_df = load_sterile_reference(default_source)
else:
    st.error(
        f"기본 엑셀 파일이 없습니다: {DEFAULT_PATH} 또는 {LIGHTWEIGHT_DEFAULT_PATH} 또는 {LIGHTWEIGHT_ASCII_PATH}"
    )
    st.stop()
sterile_order_map = _build_sterile_order_map(sterile_ref_df)
sterile_ref_view_df = _build_sterile_reference_view(sterile_ref_df, before_count=20, after_count=20)

with meta_col:
    if inventory_upload is not None:
        st.caption(f"재고장 업데이트: 업로드 파일 ({inventory_upload.name})")
    elif source_path is not None and Path(source_path).exists():
        mtime = Path(source_path).stat().st_mtime
        kst = timezone(timedelta(hours=9))
        mtime_kst = datetime.fromtimestamp(mtime, tz=kst)
        st.caption(f"재고장 업데이트: {mtime_kst:%Y-%m-%d %H:%M} (KST)")

if show_sterile_ref:
    header_left, header_right = st.columns([4, 3])
    with header_left:
        st.subheader("멸균넘버 기준표")
    with header_right:
        f0, f1, f2 = st.columns([1.2, 2, 1])
        with f0:
            sterile_date_field = st.selectbox(
                "검색 기준",
                options=["시작일자", "유효기간(5년)", "유효기간(8년)"],
                key="sterile_ref_date_field",
                label_visibility="collapsed",
            )
        with f1:
            sterile_date_value = st.date_input(
                "시작일자 검색",
                value=None,
                key="sterile_ref_date",
                label_visibility="collapsed",
            )
        with f2:
            sterile_date_mode = st.radio(
                "검색 방향",
                options=["이후", "이전"],
                key="sterile_ref_date_mode",
                horizontal=True,
                label_visibility="collapsed",
            )

    if sterile_ref_view_df.empty:
        st.info("멸균넘버 기준표를 불러오지 못했습니다.")
    else:
        filtered_ref = sterile_ref_view_df.copy()
        date_value = None
        if sterile_date_value:
            date_value = pd.to_datetime(sterile_date_value, errors="coerce")
        if date_value is not None and not pd.isna(date_value):
            base_dates = pd.to_datetime(filtered_ref[sterile_date_field], errors="coerce")
            if sterile_date_mode == "이전":
                filtered_ref = filtered_ref[base_dates <= date_value]
            else:
                filtered_ref = filtered_ref[base_dates >= date_value]
            filtered_ref = filtered_ref.sort_values(
                sterile_date_field,
                ascending=(sterile_date_mode != "이전"),
            )

        highlight_idx = None
        if date_value is not None and not pd.isna(date_value):
            base_dates = pd.to_datetime(filtered_ref[sterile_date_field], errors="coerce")
            exact = filtered_ref[base_dates == date_value]
            if not exact.empty:
                highlight_idx = set(exact.index)
            else:
                if sterile_date_mode == "이전":
                    candidate = filtered_ref.loc[base_dates == base_dates.max()]
                else:
                    candidate = filtered_ref.loc[base_dates == base_dates.min()]
                if not candidate.empty:
                    highlight_idx = {candidate.index[0]}

        def _highlight_selected_date(row):
            if highlight_idx and row.name in highlight_idx:
                return ["background-color: #d8ecff"] * len(row)
            return [""] * len(row)

        st.dataframe(
            filtered_ref.style.apply(_highlight_selected_date, axis=1),
            use_container_width=True,
            height=720,
        )

ocr_query = ""
if show_ocr:
    st.subheader("OCR 검색")
    ocr_left, ocr_right = st.columns([2, 1])
    with ocr_left:
        ocr_upload = st.file_uploader(
            "OCR 이미지 업로드",
            type=["png", "jpg", "jpeg"],
            key="ocr_upload",
        )
    with ocr_right:
        if _paste_image_button is not None:
            paste_result = _paste_image_button("클립보드 이미지 붙여넣기", key="ocr_paste")
            if paste_result and getattr(paste_result, "image_data", None) is not None:
                st.session_state["ocr_paste_image"] = paste_result.image_data
        else:
            paste_result = None
            st.caption("붙여넣기 기능을 불러오지 못했습니다.")

    if "ocr_text" not in st.session_state:
        st.session_state["ocr_text"] = ""
    if "ocr_query" not in st.session_state:
        st.session_state["ocr_query"] = ""

    if st.button("OCR 실행"):
        image_bytes = None
        if "ocr_paste_image" in st.session_state:
            image_bytes = _image_to_bytes(st.session_state["ocr_paste_image"])
        if not image_bytes and ocr_upload is not None:
            image_bytes = ocr_upload.getvalue()

        if not image_bytes:
            st.warning("OCR에 사용할 이미지를 업로드하거나 붙여넣기 해주세요.")
        else:
            with st.spinner("OCR 처리 중..."):
                try:
                    ocr_text = run_ocr(image_bytes)
                    st.session_state["ocr_text"] = ocr_text
                    st.session_state["ocr_query"] = ocr_text.replace("\n", " ").strip()
                    if st.session_state["ocr_query"]:
                        st.success("OCR 결과를 검색어에 반영했습니다.")
                    else:
                        st.info("OCR 결과가 비어 있습니다.")
                except Exception as exc:
                    st.error(f"OCR 실패: {exc}")

    ocr_query = st.text_input(
        "OCR 검색어",
        key="ocr_query",
        placeholder="이미지에서 추출된 텍스트가 여기에 들어갑니다",
    )
    st.text_area("OCR 원문", key="ocr_text", height=120)

if show_search:
    header_left, header_right = st.columns([1, 7])
    with header_left:
        st.subheader("통합 검색")
    with header_right:
        # keep tight vertical spacing, but align sterile filter widths to:
        # query = AXIS~ADD (2 cols), mode = category (1 col)
        h1, h2, h3, h4 = st.columns([0.9, 3.1, 2.0, 1.0])
        with h1:
            toric_mf = st.checkbox("Toric+M/F", key="toric_mf")
        with h2:
            export_no_power = st.checkbox("Power Off", key="export_no_power")
        with h3:
            st.markdown("**멸균필터**")
            sterile_query = st.text_input(
                "멸균필터-코드",
                placeholder="예: OG83",
                key="sterile_query",
                label_visibility="collapsed",
            )
        with h4:
            st.markdown("&nbsp;", unsafe_allow_html=True)
            sterile_mode = st.selectbox(
                "멸균필터-조건",
                options=["고정", "이상", "이하"],
                key="sterile_mode",
                label_visibility="collapsed",
            )

    # Row: 코드~분류를 한 줄에 정렬
    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns([2, 1, 1, 1, 1, 1, 1, 1])
    with col1:
        query = st.text_input("코드/품명", placeholder="예: T4556, P4050, NewFusion", key="query")
    with col2:
        color_query = st.text_input("컬러(컬러코드)", placeholder="예: Blue, BLU", key="color_query")
    with col3:
        tone_query = st.text_input("톤수", placeholder="예: 2", key="tone_query")
    with col4:
        power_query = st.text_input("파워", placeholder="예: -02.50", key="power_query")
    with col5:
        cyl_query = st.text_input("CYL", placeholder="예: -01.25", key="cyl_query")
    with col6:
        axis_query = st.text_input("AXIS", placeholder="예: 90", key="axis_query")
    with col7:
        add_query = st.text_input("ADD", placeholder="예: +1.00", key="add_query")
    with col8:
        category_query = st.selectbox("분류", options=["전체", "FRP", "1DAY"], key="category_query")
else:
    query = st.session_state.get("query", "")
    color_query = st.session_state.get("color_query", "")
    sterile_query = st.session_state.get("sterile_query", "")
    sterile_mode = st.session_state.get("sterile_mode", "고정")
    tone_query = st.session_state.get("tone_query", "")
    power_query = st.session_state.get("power_query", "")
    cyl_query = st.session_state.get("cyl_query", "")
    axis_query = st.session_state.get("axis_query", "")
    add_query = st.session_state.get("add_query", "")
    toric_mf = st.session_state.get("toric_mf", False)
    export_no_power = st.session_state.get("export_no_power", False)
    category_query = st.session_state.get("category_query", "전체")


missing_cols: list[str] = []


def _filter_contains(df: pd.DataFrame, col: str, value: str) -> pd.DataFrame:
    if not value:
        return df
    if col not in df.columns:
        missing_cols.append(col)
        return df
    parts = [v.strip().upper() for v in value.split(",") if v.strip()]
    if not parts:
        return df
    series = df[col].fillna("").astype(str).str.strip().str.upper()
    mask = series.str.contains(parts[0], na=False)
    for p in parts[1:]:
        mask = mask | series.str.contains(p, na=False)
    return df[mask]


def _filter_numeric_equal(df: pd.DataFrame, col: str, value: str, decimals: int = 2) -> pd.DataFrame:
    if not value:
        return df
    if col not in df.columns:
        missing_cols.append(col)
        return df
    parts = [v.strip() for v in value.split(",") if v.strip()]
    if not parts:
        return df
    targets = []
    for p in parts:
        try:
            targets.append(round(float(p), decimals))
        except Exception:
            continue
    if not targets:
        return df
    series = pd.to_numeric(df[col], errors="coerce").round(decimals)
    mask = False
    for t in targets:
        mask = mask | (series == t)
    return df[mask]


def _filter_int_equal(df: pd.DataFrame, col: str, value: str) -> pd.DataFrame:
    if not value:
        return df
    if col not in df.columns:
        missing_cols.append(col)
        return df
    parts = [v.strip() for v in value.split(",") if v.strip()]
    if not parts:
        return df
    targets = []
    for p in parts:
        try:
            targets.append(int(p))
        except Exception:
            continue
    if not targets:
        return df
    series = pd.to_numeric(df[col], errors="coerce").round(0)
    mask = False
    for t in targets:
        mask = mask | (series == t)
    return df[mask]


def _filter_color(df: pd.DataFrame, value: str) -> pd.DataFrame:
    if not value:
        return df
    parts = [v.strip().upper().replace(" ", "") for v in value.split(",") if v.strip()]
    if not parts:
        return df
    masks = []
    if COL_COLOR in df.columns:
        s = df[COL_COLOR].fillna("").astype(str).str.upper().str.replace(r"\s+", "", regex=True)
        m = s.str.contains(parts[0], na=False)
        for p in parts[1:]:
            m = m | s.str.contains(p, na=False)
        masks.append(m)
    if "컬러코드" in df.columns:
        s = df["컬러코드"].fillna("").astype(str).str.upper().str.replace(r"\s+", "", regex=True)
        m = s.str.contains(parts[0], na=False)
        for p in parts[1:]:
            m = m | s.str.contains(p, na=False)
        masks.append(m)
    if not masks:
        missing_cols.append("컬러/컬러코드")
        return df
    mask = masks[0]
    for m in masks[1:]:
        mask = mask | m
    return df[mask]


def _filter_sterile_number(df: pd.DataFrame, value: str, mode: str, order_map: dict[str, int]) -> pd.DataFrame:
    target = str(value or "").strip().upper()
    if not target:
        return df
    sterile_col = None
    for c in ["멸균No.", "멸균NO.", "멸균no."]:
        if c in df.columns:
            sterile_col = c
            break
    if sterile_col is None:
        missing_cols.append("멸균No.")
        return df
    lots = df[sterile_col].fillna("").astype(str).str.strip().str.upper()
    if mode == "고정":
        return df[lots == target]

    target_idx = _sterile_code_to_rank(target)
    if target_idx is None and order_map and target in order_map:
        target_idx = order_map[target]

    idx = lots.map(_sterile_code_to_rank)
    if order_map:
        idx = idx.fillna(lots.map(order_map))

    if target_idx is not None:
        if mode == "이상":
            return df[idx.notna() & (idx >= target_idx)]
        if mode == "이하":
            return df[idx.notna() & (idx <= target_idx)]

    # 기준표에 없는 코드이거나 기준표를 못 읽었을 때는 문자열 비교로 fallback
    if mode == "이상":
        return df[lots >= target]
    if mode == "이하":
        return df[lots <= target]
    return df[lots == target]


has_filters = any(
    [
        query.strip(),
        ocr_query.strip(),
        color_query.strip(),
        sterile_query.strip(),
        tone_query.strip(),
        power_query.strip(),
        cyl_query.strip(),
        axis_query.strip(),
        add_query.strip(),
    ]
)

if show_search and has_filters:
    filtered_df = df
    filtered_df = _filter_color(filtered_df, color_query)
    filtered_df = _filter_sterile_number(filtered_df, sterile_query, sterile_mode, sterile_order_map)
    filtered_df = _filter_contains(filtered_df, COL_TONE, tone_query)
    filtered_df = _filter_numeric_equal(filtered_df, COL_POWER, power_query, decimals=2)
    filtered_df = _filter_numeric_equal(filtered_df, COL_CYL, cyl_query, decimals=2)
    filtered_df = _filter_int_equal(filtered_df, COL_AXIS, axis_query)
    filtered_df = _filter_numeric_equal(filtered_df, COL_ADD, add_query, decimals=2)
    if category_query != "전체":
        filtered_df = _filter_contains(filtered_df, "분류", category_query)

    if missing_cols:
        missing_label = ", ".join(sorted(set(missing_cols)))
        st.warning(f"엑셀에 다음 컬럼이 없어 해당 조건은 무시했습니다: {missing_label}")

    combined_query = " ".join([query.strip(), ocr_query.strip()]).strip()
    if combined_query:
        result = search_inventory(filtered_df, combined_query)
    else:
        result = summarize_inventory(filtered_df)
    if result.empty:
        st.info("검색 결과가 없습니다.")
    else:
        display_df = result.copy()
        if COL_ITEM in display_df.columns and COL_NAME in display_df.columns:
            cols = list(display_df.columns)
            cols = [c for c in cols if c != COL_ITEM]
            name_idx = cols.index(COL_NAME)
            cols.insert(name_idx, COL_ITEM)
            display_df = display_df[cols]
        if COL_COLOR_HEX in display_df.columns:
            display_df = display_df.drop(columns=[COL_COLOR_HEX])
        if "최종재고" in display_df.columns:
            display_df = display_df.rename(columns={"최종재고": "재고(pcs)"})
        if "재고(pcs)" in display_df.columns:
            display_df["재고(pcs)"] = display_df["재고(pcs)"].apply(
                lambda v: f"{int(v):,}" if pd.notna(v) and str(v).strip() != "" else v
            )

        left_col, right_col = st.columns([3, 1])
        with left_col:
            table = st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row",
                height=460,
            )

        with right_col:
            if table and table.selection and table.selection.rows:
                sel_idx = table.selection.rows[0]
                selected_row = result.reset_index(drop=True).loc[sel_idx].to_dict()
                lot_df = lot_breakdown(df, selected_row)
                st.caption("LOT 상세")
                if not lot_df.empty:
                    if "최종재고" in lot_df.columns:
                        lot_df = lot_df.rename(columns={"최종재고": "재고(pcs)"})
                    if "재고(pcs)" in lot_df.columns:
                        lot_df["재고(pcs)"] = lot_df["재고(pcs)"].apply(
                            lambda v: f"{int(v):,}" if pd.notna(v) and str(v).strip() != "" else v
                        )
                    st.dataframe(lot_df, use_container_width=True, hide_index=True, height=320)
                else:
                    st.caption("LOT 정보 없음")

        export_df = result.reset_index(drop=True)
        if not export_df.empty:
            export_kwargs = dict(
                rows=export_df.to_dict(orient="records"),
                source_path=source_path,
                source_bytes=source_bytes,
                use_toric=toric_mf,
            )
            try:
                export_bytes, not_found = build_summary_export_multi(
                    **export_kwargs,
                    remove_powers=export_no_power,
                )
            except TypeError as e:
                # Backward compatibility: older deployed module may not support remove_powers.
                if "remove_powers" not in str(e):
                    raise
                export_bytes, not_found = build_summary_export_multi(**export_kwargs)
            if not_found:
                st.warning(f"SUMMARY 시트에서 {not_found}개 품목을 찾지 못해 수량이 0으로 출력될 수 있습니다.")

            file_name = "SEARCH_RESULT_SUMMARY_EXPORT.xlsx"

            left, _ = st.columns([1, 5])
            with left:
                st.download_button(
                    "엑셀로 내보내기",
                    data=export_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
elif show_search:
    st.info("검색 조건을 입력하면 결과가 표시됩니다.")
elif show_ocr:
    st.info("OCR 화면입니다. OCR 실행 후 검색어를 확인하세요.")
elif show_sterile_ref:
    pass
else:
    st.info("왼쪽 사이드바에서 검색 화면을 선택하세요.")


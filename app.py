import streamlit as st
from pathlib import Path
from datetime import datetime, timezone, timedelta
from io import BytesIO
import json
import pandas as pd

from inventory_search import load_inventory, load_inventory_from_bytes, search_inventory, summarize_inventory, lot_breakdown
from inventory_search import (
    build_summary_export_multi,
    COL_COLOR,
    COL_COLOR_HEX,
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

try:
    from streamlit_paste_button import paste_image_button as _paste_image_button
except Exception:
    _paste_image_button = None

st.set_page_config(page_title="재고 검색", layout="wide")

title_col, meta_col = st.columns([3, 1])
with title_col:
    st.title("재고 검색 대시보드")

with st.sidebar:
    st.header("데이터")
    st.caption("기본 경로: " + str(DEFAULT_PATH))
    st.header("검색 표시")
    show_ocr = st.checkbox("OCR 검색", value=False)
    show_search = st.checkbox("통합 검색", value=True)

@st.cache_data(show_spinner=False)
def load_from_path(path: Path) -> pd.DataFrame:
    return load_inventory(path, load_color=False)

@st.cache_data(show_spinner=False)
def load_from_bytes(data: bytes) -> pd.DataFrame:
    return load_inventory_from_bytes(data, load_color=False)


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
if not DEFAULT_PATH.exists():
    st.error(f"기본 엑셀 파일이 없습니다: {DEFAULT_PATH}")
    st.stop()
df = load_from_path(DEFAULT_PATH)
source_path = DEFAULT_PATH

with meta_col:
    if source_path is not None and Path(source_path).exists():
        mtime = Path(source_path).stat().st_mtime
        kst = timezone(timedelta(hours=9))
        mtime_kst = datetime.fromtimestamp(mtime, tz=kst)
        st.caption(f"재고장 업데이트: {mtime_kst:%Y-%m-%d %H:%M} (KST)")

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
    header_left, header_right = st.columns([1, 6])
    with header_left:
        st.subheader("통합 검색")
    with header_right:
        toric_mf = st.checkbox("Toric+M/F", value=False)
    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns([2, 1, 1, 1, 1, 1, 1, 1])
    with col1:
        query = st.text_input("코드/품명", placeholder="예: T4556, P4050, NewFusion", key="query")
    with col2:
        color_query = st.text_input("컬러(컬러코드)", placeholder="예: Blue, BLU")
    with col3:
        tone_query = st.text_input("톤수", placeholder="예: 2")
    with col4:
        power_query = st.text_input("파워", placeholder="예: -02.50")
    with col5:
        cyl_query = st.text_input("CYL", placeholder="예: -01.25")
    with col6:
        axis_query = st.text_input("AXIS", placeholder="예: 90")
    with col7:
        add_query = st.text_input("ADD", placeholder="예: +1.00")
    with col8:
        category_query = st.selectbox("분류", options=["전체", "FRP", "1DAY"], index=0)
else:
    query = ""
    color_query = ""
    tone_query = ""
    power_query = ""
    cyl_query = ""
    axis_query = ""
    add_query = ""
    toric_mf = False
    category_query = "전체"


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

has_filters = any(
    [
        query.strip(),
        ocr_query.strip(),
        color_query.strip(),
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
                    st.dataframe(lot_df, use_container_width=True, height=320)
                else:
                    st.caption("LOT 정보 없음")

        export_df = result.reset_index(drop=True)
        if not export_df.empty:
            export_bytes, not_found = build_summary_export_multi(
                rows=export_df.to_dict(orient="records"),
                source_path=source_path,
                source_bytes=source_bytes,
                use_toric=toric_mf,
            )
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
else:
    st.info("왼쪽 사이드바에서 검색 화면을 선택하세요.")

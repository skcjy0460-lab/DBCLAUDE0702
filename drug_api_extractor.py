# -*- coding: utf-8 -*-
"""
ClaimLens 부속 도구 - 공공데이터포털 의약품 API 통합 추출기
====================================================================
목적
  1) 품목 단위 상세조회(개별검색) : 약품 하나를 검색하면 4개 공공API의
     정보를 한 화면에서 카드 형태로 확인
  2) 전체 리스트 대량추출(벌크수집) : 조건(제조사/허가일/전문·일반 등)에
     맞는 데이터를 API가 끝날 때까지 페이지네이션을 자동으로 반복 수집하여
     하나의 DataFrame으로 합친 뒤 엑셀(.xlsx)로 저장 → 자체 SQLite DB 적재용
     원본 자료로 사용

대상 API (모두 apis.data.go.kr 도메인, 공공데이터포털 활용신청 필요)
  A. 건강보험심사평가원_약가기준정보조회서비스 (약가마스터)
     base : http://apis.data.go.kr/B551182/dgamtCrtrInfoService1.2
     오퍼레이션 : getDgamtList  ← data.go.kr 활용신청 상세페이지에서 확인됨(확정)
  B. 식품의약품안전처_의약품 제품 허가정보서비스
     base : http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07
     오퍼레이션(목록) : getDrugPrdtPrmsnInq07 ← 확정
     오퍼레이션(상세/효능효과·용법용량·주의사항) : getDrugPrdtPrmsnDtlInq06 ← 확정 (주의: 버전번호가 06으로 목록조회의 07과 다름)
     오퍼레이션(주성분 상세) : getDrugPrdtMcpnDtlInq07 ← 확정
  C. 식품의약품안전처_DUR품목정보서비스 (병용금기/임부금기/연령금기/노인주의 등)
     base : http://apis.data.go.kr/1471000/DURPrdlstInfoService
     오퍼레이션(병용금기) : getUsjntTabooInfoList ← 확정
     그 외 오퍼레이션(임부금기/연령금기/노인주의 등)은 예상값이며 주석 참고
  D. 식품의약품안전처_의약품개요정보(e약은요) - 일반의약품만 커버(참고용)
     base : http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService02 (예상값)

※ "확정"이 아닌 값은 본인 마이페이지 > 활용신청 상세 > "요청주소" 복사본으로
   반드시 1회 검증 후 아래 CONFIG 섹션만 고쳐 쓰면 됩니다. (검색-치환만 하면 됨)
"""

import io
import time
import urllib.parse as up
import xml.etree.ElementTree as ET

import pandas as pd
import requests
import streamlit as st

# ============================================================
# 0. 페이지 설정
# ============================================================
st.set_page_config(page_title="ClaimLens 약품 API 추출기", page_icon="💊", layout="wide")

# ============================================================
# 1. API 설정 (CONFIG) - 여기만 고치면 전체 프로그램에 반영됨
# ============================================================
API_CONFIG = {
    "hira_price": {
        "label": "① 심평원 약가마스터 (보험코드/상한가/급여여부)",
        "base_url": "http://apis.data.go.kr/B551182/dgamtCrtrInfoService1.2",
        "operation": "getDgamtList",
        "confirmed": True,
        "search_param_candidates": ["itemNm", "prdtNm"],  # 제품명 검색 파라미터 후보(응답 확인 후 확정 권장)
        "date_param": None,
        "fields_of_interest": {
            "gnlNmCd": "일반명코드",
            "prdtNm": "제품명",
            "entpNm": "업체명",
            "spGnlNmCd": "보험코드(주성분코드)",
            "meftDivNo": "약효분류번호",
            "chgDate": "고시(변경)일자",
            "amt": "상한금액",
            "frmlyDivNm": "제형구분명",
        },
    },
    "mfds_permit_list": {
        "label": "② 식약처 의약품 제품 허가정보 (제품명/성분/제형/허가일/전문·일반)",
        "base_url": "http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07",
        "operation": "getDrugPrdtPrmsnInq07",
        "confirmed": True,
        "search_param_candidates": ["item_name", "entp_name"],
        "date_param": ("item_permit_date", "허가일자(YYYYMMDD)"),
        "fields_of_interest": {
            "ITEM_SEQ": "품목기준코드",
            "ITEM_NAME": "제품명",
            "ENTP_NAME": "업체명",
            "ITEM_PERMIT_DATE": "허가일자",
            "ETC_OTC_NAME": "전문/일반",
            "CHART": "성상",
            "MATERIAL_NAME": "성분(원료성분)",
            "CANCEL_NAME": "취소상태",
            "STORAGE_METHOD": "저장방법",
            "PACK_UNIT": "포장단위",
        },
    },
    "mfds_permit_detail": {
        "label": "③ 식약처 의약품 상세정보 (효능효과/용법용량/사용상주의사항)",
        "base_url": "http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07",
        "operation": "getDrugPrdtPrmsnDtlInq06",  # 조정윤님 마이페이지 Swagger에서 확인됨 (목록조회와 버전번호가 다름: 06)
        "confirmed": True,  # ← 2026-07-02 사용자 계정 Swagger 화면으로 확인 완료
        "search_param_candidates": ["item_seq", "item_name"],
        "date_param": None,
        "fields_of_interest": {
            "ITEM_SEQ": "품목기준코드",
            "ITEM_NAME": "제품명",
            "EE_DOC_DATA": "효능효과",
            "UD_DOC_DATA": "용법용량",
            "NB_DOC_DATA": "사용상주의사항",
        },
    },
    "mfds_permit_ingredient_detail": {
        "label": "③-보조 식약처 의약품 주성분 상세정보",
        "base_url": "http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07",
        "operation": "getDrugPrdtMcpnDtlInq07",  # 조정윤님 마이페이지 Swagger에서 확인됨
        "confirmed": True,
        "search_param_candidates": ["item_seq", "item_name"],
        "date_param": None,
        "fields_of_interest": {
            "ITEM_SEQ": "품목기준코드",
            "ITEM_NAME": "제품명",
            "MATERIAL_NAME": "주성분명",
            "MAIN_INGR_ENG": "주성분 영문명",
            "TOTAL_CONTENT": "함량",
            "UNIT": "단위",
        },
    },
    "dur_taboo": {
        "label": "④ DUR 병용금기 (Interaction)",
        "base_url": "http://apis.data.go.kr/1471000/DURPrdlstInfoService",
        "operation": "getUsjntTabooInfoList",
        "confirmed": True,
        "search_param_candidates": ["ITEM_NAME"],
        "date_param": None,
        "fields_of_interest": {
            "ITEM_SEQ": "품목기준코드",
            "ITEM_NAME": "제품명",
            "MIXTURE_ITEM_SEQ": "병용금기 상대 품목기준코드",
            "MIXTURE_ITEM_NAME": "병용금기 상대 제품명",
            "PROHBT_CONTENT": "금기내용",
            "NOTIFICATION_DATE": "고시일자",
        },
    },
    "dur_pregnant": {
        "label": "④ DUR 임부금기",
        "base_url": "http://apis.data.go.kr/1471000/DURPrdlstInfoService",
        "operation": "getPwnmTabooInfoList",  # 예상값 - 마이페이지 확인 권장
        "confirmed": False,
        "search_param_candidates": ["ITEM_NAME"],
        "date_param": None,
        "fields_of_interest": {
            "ITEM_SEQ": "품목기준코드",
            "ITEM_NAME": "제품명",
            "PROHBT_CONTENT": "금기내용/등급",
            "NOTIFICATION_DATE": "고시일자",
        },
    },
    "dur_age": {
        "label": "④ DUR 연령금기 (특정연령대금기)",
        "base_url": "http://apis.data.go.kr/1471000/DURPrdlstInfoService",
        "operation": "getSpcifyAgrdeTabooInfoList",  # 예상값 - 마이페이지 확인 권장
        "confirmed": False,
        "search_param_candidates": ["ITEM_NAME"],
        "date_param": None,
        "fields_of_interest": {
            "ITEM_SEQ": "품목기준코드",
            "ITEM_NAME": "제품명",
            "PROHBT_CONTENT": "금기내용(연령기준)",
            "NOTIFICATION_DATE": "고시일자",
        },
    },
    "dur_elderly": {
        "label": "④ DUR 노인주의",
        "base_url": "http://apis.data.go.kr/1471000/DURPrdlstInfoService",
        "operation": "getOdsnAtentInfoList",  # 예상값 - 마이페이지 확인 권장
        "confirmed": False,
        "search_param_candidates": ["ITEM_NAME"],
        "date_param": None,
        "fields_of_interest": {
            "ITEM_SEQ": "품목기준코드",
            "ITEM_NAME": "제품명",
            "PROHBT_CONTENT": "주의내용",
            "NOTIFICATION_DATE": "고시일자",
        },
    },
}

# ============================================================
# 2. 공통 HTTP / 파싱 유틸
# ============================================================
def normalize_service_key(raw_key: str, key_mode: str) -> str:
    """
    공공데이터포털은 '인증키(Encoding)'와 '인증키(Decoding)' 두 종류를 발급한다.
    - Encoding 키를 그대로 requests의 params에 넣으면 requests가 다시 한 번
      URL 인코딩을 해버려서 '%25'가 섞이는 이중 인코딩(double-encoding) 오류가 난다.
    - 해결책: 항상 Decoding(raw) 형태로 맞춘 뒤 requests가 1회만 인코딩하도록 한다.
    """
    if key_mode == "인코딩(Encoding) 키를 붙여넣었어요":
        try:
            return up.unquote(raw_key)
        except Exception:
            return raw_key
    return raw_key  # 이미 Decoding 키


def build_params(service_key: str, num_rows: int, page_no: int, extra: dict) -> dict:
    params = {
        "serviceKey": service_key,
        "numOfRows": num_rows,
        "pageNo": page_no,
        "type": "json",  # 대부분 최신 오퍼레이션은 JSON 지원. 미지원 시 XML 폴백 처리함.
    }
    params.update({k: v for k, v in extra.items() if v not in (None, "")})
    return params


def parse_xml_items(xml_text: str):
    """공공데이터포털 표준 XML 응답을 dict 리스트로 변환"""
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return [], 0, "XML 파싱 실패 (응답 원문을 확인하세요)"

    header = root.find(".//header")
    result_code = header.findtext("resultCode") if header is not None else None
    result_msg = header.findtext("resultMsg") if header is not None else None

    total_count_el = root.find(".//totalCount")
    total_count = int(total_count_el.text) if total_count_el is not None and total_count_el.text else 0

    items = []
    for item in root.findall(".//item"):
        row = {child.tag: (child.text or "") for child in item}
        items.append(row)

    if result_code not in (None, "00", "0"):
        return items, total_count, f"[{result_code}] {result_msg}"
    return items, total_count, None


def call_api(base_url: str, operation: str, params: dict, timeout: int = 15):
    """
    단일 페이지 호출. JSON 우선 시도 → 실패/비JSON이면 XML 파싱으로 폴백.
    반환: (items:list[dict], total_count:int, error:str|None, raw_status:int)
    """
    url = f"{base_url}/{operation}"
    try:
        resp = requests.get(url, params=params, timeout=timeout)
    except requests.exceptions.RequestException as e:
        return [], 0, f"네트워크 오류: {e}", None

    status = resp.status_code
    text = resp.text.strip()

    # data.go.kr 공통 에러(서비스키 미등록/트래픽초과 등)는 종종 XML로만 내려온다
    if text.startswith("<"):
        items, total, err = parse_xml_items(text)
        return items, total, err, status

    # JSON 응답 처리
    try:
        data = resp.json()
    except ValueError:
        return [], 0, f"알 수 없는 응답 형식 (status={status}): {text[:200]}", status

    try:
        # 일부 API(예: DrugPrdtPrmsnInfoService07)는 "response" 겉껍질 없이
        # {"header":..., "body":...}를 최상위로 바로 내려준다. 두 구조 모두 지원.
        if "response" in data and isinstance(data["response"], dict):
            root_obj = data["response"]
        else:
            root_obj = data

        body = root_obj["body"]
        header = root_obj["header"]
        result_code = str(header.get("resultCode", ""))
        result_msg = header.get("resultMsg", "")
        total_count = int(body.get("totalCount", 0) or 0)
        items_raw = body.get("items", "")
        if items_raw in ("", None):
            items = []
        elif isinstance(items_raw, dict) and "item" in items_raw:
            item_val = items_raw["item"]
            items = item_val if isinstance(item_val, list) else [item_val]
        elif isinstance(items_raw, list):
            items = items_raw
        else:
            items = []
        err = None if result_code in ("00", "0") else f"[{result_code}] {result_msg}"
        return items, total_count, err, status
    except (KeyError, TypeError):
        return [], 0, f"예상치 못한 JSON 구조: {str(data)[:300]}", status


def fetch_all_pages(base_url, operation, service_key, extra_params, num_rows=100,
                     max_pages=500, sleep_sec=0.15, progress_cb=None,
                     start_index=None, end_index=None):
    """
    totalCount를 확인하면서 pageNo를 계속 증가시켜 데이터를 수집한다.
    progress_cb(cur_page, total_pages, collected_count) 형태의 콜백을 넘기면
    Streamlit 진행률 표시에 사용할 수 있다.

    start_index / end_index : 1부터 시작하는 전역 순번(1-based, 양끝 포함) 범위.
      지정하면 전체가 아니라 해당 구간만 수집한다.
      예) start_index=10001, end_index=20000 → 10,001번째~20,000번째 데이터만 수집.
      대량 데이터(수만 건)를 한 번에 처리하면 서버 타임아웃/메모리 문제가 생길 수 있어
      1만 건 단위 등으로 나눠 여러 번 실행할 때 사용한다.
    """
    all_items = []
    total_count = None
    errors = []

    start_page = ((start_index - 1) // num_rows + 1) if start_index and start_index > 1 else 1
    page = start_page
    pages_fetched = 0

    while True:
        params = build_params(service_key, num_rows, page, extra_params)
        items, total, err, status = call_api(base_url, operation, params)

        if err:
            errors.append(f"page {page}: {err}")
            # 인증/트래픽 오류로 보이면 즉시 중단, 그 외 단일 페이지 오류는 스킵 후 계속
            if any(kw in (err or "") for kw in ["SERVICE_KEY", "LIMITED", "22", "30", "31"]):
                break

        if total_count is None:
            total_count = total

        page_start_idx = (page - 1) * num_rows + 1
        page_end_idx = page_start_idx + len(items) - 1

        page_items = items
        if start_index or end_index:
            lo = start_index or 1
            hi = end_index if end_index else float("inf")
            page_items = [
                it for offset, it in enumerate(items)
                if lo <= (page_start_idx + offset) <= hi
            ]

        all_items.extend(page_items)
        pages_fetched += 1

        if progress_cb:
            if end_index:
                total_target = end_index - (start_index or 1) + 1
                total_pages_est = max(1, -(-total_target // num_rows))
            else:
                total_pages_est = max(1, -(-max(total_count or 0, 1) // num_rows))
            progress_cb(pages_fetched, total_pages_est, len(all_items))

        time.sleep(sleep_sec)

        no_more_data = not items
        reached_end_index = bool(end_index) and page_end_idx >= end_index
        reached_total = bool(total_count) and page_end_idx >= total_count
        page += 1

        if no_more_data or reached_end_index or reached_total:
            break
        if pages_fetched >= max_pages:
            errors.append(f"안전장치 작동: max_pages({max_pages}) 도달, 수집을 중단합니다.")
            break

    return all_items, total_count or len(all_items), errors


def to_excel_bytes(sheets: dict) -> bytes:
    """sheets = {"시트이름": DataFrame, ...} → 엑셀 바이트로 변환 (열너비 자동조정 포함)"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]  # 엑셀 시트명 31자 제한
            df.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.sheets[safe_name]
            for i, col in enumerate(df.columns, start=1):
                max_len = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str).head(200)])
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max_len + 3, 50)
            ws.freeze_panes = "A2"
    return buf.getvalue()


def rename_columns(df: pd.DataFrame, field_map: dict) -> pd.DataFrame:
    keep = [c for c in field_map if c in df.columns]
    extra_cols = [c for c in df.columns if c not in field_map]
    ordered = keep + extra_cols
    df = df[ordered].rename(columns=field_map)
    return df


# ============================================================
# 3. 사이드바 - 인증키 / 공통 옵션
# ============================================================
with st.sidebar:
    st.header("🔑 인증 설정")
    service_key_raw = st.text_input(
        "공공데이터포털 서비스키",
        type="password",
        help="마이페이지 > 개발계정 상세보기에서 발급받은 키를 붙여넣으세요.",
    )
    key_mode = st.radio(
        "붙여넣은 키의 종류",
        ["디코딩(Decoding) 키를 붙여넣었어요", "인코딩(Encoding) 키를 붙여넣었어요"],
        index=0,
        help="어떤 키인지 헷갈리면 '인코딩' 선택 시 키 끝부분에 %2B, %3D 같은 문자가 있는지 보세요. "
             "있으면 인코딩 키입니다. 이중 인코딩 오류(SERVICE KEY IS NOT REGISTERED 등)가 나면 이 옵션을 바꿔보세요.",
    )
    service_key = normalize_service_key(service_key_raw, key_mode)

    st.divider()
    st.header("⚙️ 수집 옵션")
    num_rows = st.number_input("페이지당 요청 건수(numOfRows)", min_value=10, max_value=1000, value=100, step=10)
    sleep_sec = st.slider("요청 간 대기시간(초)", 0.0, 1.0, 0.15, 0.05,
                           help="너무 빠르게 연속 요청하면 트래픽 초과 오류가 날 수 있습니다.")
    max_pages = st.number_input("최대 페이지 안전장치", min_value=10, max_value=5000, value=500, step=10)

    st.divider()
    st.caption(
        "⚠️ 표에 '예상값' 표시가 있는 오퍼레이션은 사용 전 "
        "마이페이지 > 활용신청 상세 > '요청주소' 예시로 1회 검증 후 "
        "코드 상단 CONFIG의 operation 값만 바꿔주시면 됩니다."
    )

st.title("💊 ClaimLens 약품 공공데이터 API 추출기")
st.caption("공공데이터포털 API → 자체 약품 DB 구축용 엑셀 변환 도구")

tab_single, tab_bulk, tab_guide = st.tabs(["🔍 개별 검색", "📦 전체 리스트 대량추출(엑셀)", "📖 사용 가이드"])

# ============================================================
# 4. TAB 1 - 개별 검색 (품목 단위 상세조회)
# ============================================================
with tab_single:
    st.subheader("품목 하나를 검색해서 4개 API 정보를 한 화면에서 확인")
    col1, col2 = st.columns([3, 1])
    with col1:
        query = st.text_input("제품명 (예: 타이레놀정500mg)", key="single_query")
    with col2:
        run_single = st.button("검색", type="primary", use_container_width=True, key="btn_single")

    if run_single:
        if not service_key:
            st.error("사이드바에 서비스키를 먼저 입력해주세요.")
        elif not query.strip():
            st.warning("제품명을 입력해주세요.")
        else:
            for key, cfg in API_CONFIG.items():
                with st.expander(cfg["label"], expanded=True):
                    if not cfg["confirmed"]:
                        st.caption("⚠️ 이 오퍼레이션 ID는 예상값입니다. 결과가 비어있으면 CONFIG의 operation 값을 확인하세요.")

                    param_name = cfg["search_param_candidates"][0]
                    extra = {param_name: query.strip()}
                    params = build_params(service_key, 10, 1, extra)

                    items, total, err, status = call_api(cfg["base_url"], cfg["operation"], params)

                    if err:
                        st.error(f"오류: {err}")
                        continue
                    if not items:
                        st.info("검색 결과가 없습니다. (파라미터명이 다를 수 있습니다 — 상세 가이드 탭 참고)")
                        continue

                    df = pd.DataFrame(items)
                    df_display = rename_columns(df, cfg["fields_of_interest"])
                    st.write(f"총 {total}건 중 {len(items)}건 표시")
                    st.dataframe(df_display, use_container_width=True, hide_index=True)

# ============================================================
# 5. TAB 2 - 전체 리스트 대량추출
# ============================================================
with tab_bulk:
    st.subheader("조건에 맞는 데이터를 끝까지 자동 수집 → 엑셀로 저장")

    source_key = st.selectbox(
        "추출할 API 선택",
        options=list(API_CONFIG.keys()),
        format_func=lambda k: API_CONFIG[k]["label"],
        key="bulk_source",
    )
    cfg = API_CONFIG[source_key]

    if not cfg["confirmed"]:
        st.warning("⚠️ 이 오퍼레이션 ID는 예상값입니다. 대량 수집 전 개별 검색 탭에서 결과가 정상적으로 나오는지 먼저 확인하세요.")

    st.markdown("**검색/필터 조건** (비워두면 조건 없이 전체 수집을 시도합니다 — API에 따라 조건 필수인 경우가 있어요)")
    fcol1, fcol2 = st.columns(2)
    with fcol1:
        filter_value = st.text_input(
            f"검색어 ({' / '.join(cfg['search_param_candidates'])} 중 첫번째 파라미터 사용)",
            key="bulk_filter_value",
        )
    with fcol2:
        date_value = ""
        if cfg["date_param"]:
            date_field, date_label = cfg["date_param"]
            date_value = st.text_input(f"{date_label}", key="bulk_date_value")

    st.markdown(
        "**🔢 대량 데이터 분할 수집** — 데이터가 2~3만 건 이상으로 많으면 한 번에 다 수집하다가 "
        "서버 타임아웃/메모리 문제로 실패할 수 있습니다. 아래에서 구간을 나눠(예: 1~10000, "
        "10001~20000 …) 여러 번 실행하면, 결과가 자동으로 이어붙여져서 마지막에 전체를 "
        "한 번에 다운로드할 수 있습니다."
    )
    use_range = st.checkbox("범위를 지정해서 수집하기", key="bulk_use_range")
    start_idx_input, end_idx_input = None, None
    if use_range:
        rcol1, rcol2 = st.columns(2)
        with rcol1:
            start_idx_input = st.number_input(
                "시작 번호", min_value=1, value=1, step=10000, key="bulk_start_idx",
                help="1부터 시작하는 순번입니다. 예: 첫 구간은 1",
            )
        with rcol2:
            end_idx_input = st.number_input(
                "종료 번호", min_value=1, value=10000, step=10000, key="bulk_end_idx",
                help="이 번호까지 포함해서 수집합니다. 예: 첫 구간은 10000, 다음 구간은 10001~20000",
            )
        if end_idx_input < start_idx_input:
            st.error("종료 번호가 시작 번호보다 작습니다. 값을 확인해주세요.")

    est_col1, est_col2 = st.columns([1, 1])
    with est_col1:
        run_bulk = st.button("🚀 수집 시작", type="primary", use_container_width=True, key="btn_bulk")
    with est_col2:
        st.caption("대량 수집은 데이터 건수에 따라 수 분 걸릴 수 있습니다.")

    if run_bulk:
        if not service_key:
            st.error("사이드바에 서비스키를 먼저 입력해주세요.")
        else:
            extra = {}
            if filter_value.strip():
                extra[cfg["search_param_candidates"][0]] = filter_value.strip()
            if cfg["date_param"] and date_value.strip():
                extra[cfg["date_param"][0]] = date_value.strip()

            progress_bar = st.progress(0.0, text="수집 준비 중...")
            status_text = st.empty()

            def progress_cb(cur_page, total_pages_est, collected):
                pct = min(cur_page / max(total_pages_est, 1), 1.0)
                progress_bar.progress(pct, text=f"{cur_page}페이지 수집 중... (누적 {collected:,}건)")

            items, total_count, errors = fetch_all_pages(
                cfg["base_url"], cfg["operation"], service_key, extra,
                num_rows=int(num_rows), max_pages=int(max_pages),
                sleep_sec=float(sleep_sec), progress_cb=progress_cb,
                start_index=int(start_idx_input) if use_range else None,
                end_index=int(end_idx_input) if use_range else None,
            )
            progress_bar.progress(1.0, text="수집 완료")

            if errors:
                with st.expander(f"⚠️ 수집 중 발생한 메시지 {len(errors)}건", expanded=False):
                    for e in errors:
                        st.text(e)

            if not items:
                st.error("수집된 데이터가 없습니다. 검색조건 또는 오퍼레이션 ID를 확인해주세요.")
            else:
                df = pd.DataFrame(items)
                df_display = rename_columns(df, cfg["fields_of_interest"])
                if use_range:
                    st.success(
                        f"[{int(start_idx_input):,}~{int(end_idx_input):,} 구간] "
                        f"전체 {total_count:,}건 중 이번 구간 {len(df_display):,}건 수집 완료"
                    )
                else:
                    st.success(f"총 {total_count:,}건 중 {len(df_display):,}건 수집 완료")
                st.dataframe(df_display.head(500), use_container_width=True, hide_index=True)
                if len(df_display) > 500:
                    st.caption(f"미리보기는 상위 500건만 표시됩니다. 전체 {len(df_display):,}건은 엑셀 다운로드로 확인하세요.")

                st.session_state["bulk_last_df"] = df_display
                st.session_state["bulk_last_source"] = cfg["label"]

                # 범위 지정 모드일 때는 구간별 결과를 자동으로 이어붙여서 누적 저장
                if use_range:
                    acc_key = f"bulk_accum_{source_key}"
                    prev = st.session_state.get(acc_key)
                    if prev is not None:
                        combined = pd.concat([prev, df_display], ignore_index=True)
                        dedup_col = None
                        for cand in ("품목기준코드", "ITEM_SEQ"):
                            if cand in combined.columns:
                                dedup_col = cand
                                break
                        if dedup_col:
                            combined = combined.drop_duplicates(subset=[dedup_col])
                        st.session_state[acc_key] = combined
                    else:
                        st.session_state[acc_key] = df_display

    # 다운로드 영역 (수집 결과가 세션에 있으면 항상 노출)
    if "bulk_last_df" in st.session_state:
        st.divider()
        st.markdown(f"**최근 수집 결과:** {st.session_state.get('bulk_last_source', '')} "
                    f"({len(st.session_state['bulk_last_df']):,}건)")
        excel_bytes = to_excel_bytes({"수집결과": st.session_state["bulk_last_df"]})
        st.download_button(
            "⬇️ 이번 구간 엑셀 다운로드",
            data=excel_bytes,
            file_name=f"{source_key}_추출결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # 범위 지정 모드로 여러 구간을 나눠 수집했을 때 → 누적(전체 합본) 결과 다운로드
    accum_key = f"bulk_accum_{source_key}"
    if accum_key in st.session_state:
        st.divider()
        st.markdown(
            f"**📚 누적 수집 결과 (지금까지 나눠서 수집한 구간을 모두 합친 데이터):** "
            f"{len(st.session_state[accum_key]):,}건"
        )
        st.caption("같은 API에서 구간(1~10000, 10001~20000 …)을 나눠 여러 번 수집하면 여기에 계속 합쳐집니다.")
        accum_bytes = to_excel_bytes({"누적결과": st.session_state[accum_key]})
        acol1, acol2 = st.columns([3, 1])
        with acol1:
            st.download_button(
                "⬇️ 누적 전체 엑셀 다운로드",
                data=accum_bytes,
                file_name=f"{source_key}_누적전체.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"btn_download_accum_{source_key}",
            )
        with acol2:
            if st.button("초기화", key=f"btn_reset_accum_{source_key}", use_container_width=True):
                del st.session_state[accum_key]
                st.rerun()

    st.divider()
    st.subheader("🔗 여러 API 결과를 품목기준코드로 병합해서 하나의 DB 원본표 만들기")
    st.caption(
        "위에서 API별로 각각 수집한 뒤, 아래에서 파일을 업로드하면 "
        "품목기준코드(ITEM_SEQ) 기준으로 좌우 병합(LEFT JOIN)한 통합표를 만들어 드립니다."
    )
    merge_files = st.file_uploader(
        "병합할 엑셀 파일들 (2개 이상, 각 파일에 ITEM_SEQ 또는 품목기준코드 컬럼 필요)",
        type=["xlsx"], accept_multiple_files=True, key="merge_uploader",
    )
    if merge_files and len(merge_files) >= 2:
        if st.button("병합 실행", key="btn_merge"):
            dfs = []
            for f in merge_files:
                d = pd.read_excel(f)
                key_col = "품목기준코드" if "품목기준코드" in d.columns else ("ITEM_SEQ" if "ITEM_SEQ" in d.columns else None)
                if key_col is None:
                    st.error(f"'{f.name}' 파일에 ITEM_SEQ 또는 품목기준코드 컬럼이 없어 병합에서 제외합니다.")
                    continue
                if key_col != "품목기준코드":
                    d = d.rename(columns={key_col: "품목기준코드"})
                d["품목기준코드"] = d["품목기준코드"].astype(str).str.strip()
                dfs.append(d)

            if len(dfs) >= 2:
                merged = dfs[0]
                for d in dfs[1:]:
                    merged = merged.merge(d, on="품목기준코드", how="left", suffixes=("", "_dup"))
                st.success(f"병합 완료: {len(merged):,}행, {len(merged.columns)}열")
                st.dataframe(merged.head(300), use_container_width=True, hide_index=True)
                merged_bytes = to_excel_bytes({"통합DB": merged})
                st.download_button(
                    "⬇️ 통합 엑셀 다운로드", data=merged_bytes,
                    file_name="약품_통합DB.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ============================================================
# 6. TAB 3 - 사용 가이드
# ============================================================
with tab_guide:
    st.subheader("사용 전 체크리스트")
    st.markdown(
        """
1. **서비스키 발급**: data.go.kr 마이페이지 > 활용신청 목록에서 아래 4개 서비스를 각각 신청
   - 건강보험심사평가원_약가기준정보조회서비스
   - 식품의약품안전처_의약품 제품 허가정보서비스
   - 식품의약품안전처_DUR품목정보서비스
   - (선택) 식품의약품안전처_의약품개요정보(e약은요) — 일반의약품만 커버되니 참고용으로만 사용

2. **이중 인코딩(double-encoding) 오류 대처법**
   - 오류 메시지에 `SERVICE_KEY_IS_NOT_REGISTERED_ERROR`가 뜨는데 키가 맞는 게 확실하다면 대부분 이 문제입니다.
   - 사이드바에서 "인코딩 키를 붙여넣었어요"로 전환해보세요. 이 앱은 내부적으로 키를 디코딩한 뒤
     `requests`가 1회만 인코딩하도록 처리합니다.

3. **파라미터명이 API마다 다릅니다**
   - 심평원 약가마스터는 `itemNm`/`prdtNm` 계열, 식약처 계열은 `item_name`(소문자+언더바) 또는
     `ITEM_NAME`(대문자, DUR)처럼 표기 규칙이 다릅니다.
   - 개별 검색 탭에서 결과가 비어 있다면, 마이페이지 활용신청 상세의 "요청메시지 명세"에서
     정확한 파라미터명을 확인해 CONFIG의 `search_param_candidates`를 고쳐주세요.

4. **e약은요의 한계**
   - e약은요는 공급실적이 있는 **일반의약품 위주**로만 제공됩니다. 전문의약품 효능효과/용법용량은
     ③ 식약처 의약품 상세정보(DrugPrdtPrmsnDtlInq) 쪽에서 가져오는 것이 원칙입니다.

5. **대량 수집 시 주의**
   - 개발계정은 트래픽 한도가 낮을 수 있습니다(서비스별로 1일 1,000~100,000건 등 상이).
   - 21,000여 건 규모의 전체 DB를 새로 구축할 때는 검색어 없이 numOfRows를 100~500 정도로
     설정하고 여러 날에 나눠 수집하는 것을 권장합니다.
        """
    )

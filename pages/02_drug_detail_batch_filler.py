# -*- coding: utf-8 -*-
"""
ClaimLens 부속 도구 - 효능효과/용법용량 ITEM_SEQ 일괄순회 채우기
====================================================================
배경
  기존 drug_api_extractor.py의 "③ 식약처 의약품 상세정보
  (getDrugPrdtPrmsnDtlInq06)"는 CONFIG(필드매핑)는 정확하지만,
  이 오퍼레이션은 원래 품목 1건(ITEM_SEQ 또는 ITEM_NAME) 단위로
  조회하는 API라서 대량추출 탭의 "검색어 1개" 방식으로는
  21,878건 전체를 채울 수 없었음 (96%가 비어있던 근본 원인).

이 페이지가 하는 일
  1) 기존 Master 엑셀(품목기준코드 컬럼 필수)을 업로드
  2) 효능효과/용법용량이 비어있는 행만 추출
  3) 지정한 범위(start_idx~end_idx)만큼 ITEM_SEQ를 하나씩 순회하며
     ③ 상세조회 API 호출 → EE_DOC_DATA/UD_DOC_DATA/NB_DOC_DATA 수집
  4) HTML 태그 제거 후 session_state에 누적 (품목기준코드 기준 중복제거)
  5) 누적분을 원본 Master와 병합해서 엑셀 다운로드
  6) 그래도 못 채운 행은 "동일 성분+함량+제형" 그룹 내 값 상속으로 2차 보충
     (신뢰도 컬럼에 "성분상속"이라고 표시해 원본 API 값과 구분)

Streamlit Cloud 타임아웃을 피하려면 한 번에 500~1000건씩 여러 세션에
나눠 돌리는 것을 권장 (사이드바에서 범위 지정).
"""

import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="효능효과/용법용량 일괄채우기", page_icon="🧩", layout="wide")

DETAIL_BASE_URL = "http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07"
DETAIL_OPERATION = "getDrugPrdtPrmsnDtlInq06"  # ③ 상세조회 (drug_api_extractor.py와 동일, 확정된 오퍼레이션)
LIST_OPERATION = "getDrugPrdtPrmsnInq07"  # ② 허가정보 목록조회 (제품명 -> 진짜 ITEM_SEQ 찾기용)

_session = requests.Session()
_adapter = requests.adapters.HTTPAdapter(pool_connections=20, pool_maxsize=20)
_session.mount("http://", _adapter)
_session.mount("https://", _adapter)


def resolve_item_seq(service_key: str, item_name: str, timeout: int = 15):
    """제품명으로 ② 허가정보 목록조회를 호출해서 식약처 진짜 ITEM_SEQ를 찾는다.
    Master 파일의 코드 컬럼이 식약처 ITEM_SEQ가 아닌 다른 코드체계(예: 심평원 계열)일 때 필수 단계."""
    url = f"{DETAIL_BASE_URL}/{LIST_OPERATION}"
    params = {
        "serviceKey": service_key,
        "item_name": item_name,
        "numOfRows": 10,
        "pageNo": 1,
        "type": "json",
    }
    try:
        resp = _session.get(url, params=params, timeout=timeout)
    except requests.exceptions.RequestException as e:
        return [], f"네트워크 오류: {e}"

    text = resp.text.strip()
    if text.startswith("<"):
        return [], f"XML 응답(에러 가능성): {text[:200]}"
    try:
        data = resp.json()
    except ValueError:
        return [], f"알 수 없는 응답: {text[:200]}"

    try:
        root_obj = data["response"] if "response" in data and isinstance(data["response"], dict) else data
        body = root_obj["body"]
        header = root_obj["header"]
        result_code = str(header.get("resultCode", ""))
        if result_code not in ("00", "0"):
            return [], f"[{result_code}] {header.get('resultMsg', '')}"
        items_raw = body.get("items", "")
        if items_raw in ("", None):
            return [], "검색 결과 없음"
        if isinstance(items_raw, dict) and "item" in items_raw:
            item_val = items_raw["item"]
            items = item_val if isinstance(item_val, list) else [item_val]
        elif isinstance(items_raw, list):
            items = items_raw
        else:
            return [], "예상치 못한 items 구조"
        return items, None
    except (KeyError, TypeError):
        return [], f"예상치 못한 JSON 구조: {str(data)[:300]}"


def clean_product_name(raw_name: str):
    """Master의 약품명에서 뒤에 붙은 포장단위 표기(예: '_(9.5g/95mL)')를 떼어내고
    (정제된 검색용 이름, 떼어낸 포장단위 힌트)를 반환한다.
    식약처 ITEM_NAME에는 이 포장단위가 없어서, 붙인 채로 검색하면 결과가 안 나온다."""
    if not raw_name:
        return "", ""
    raw_name = str(raw_name).strip()
    m = re.match(r"^(.*?)_?\(([^)]*(?:g|mL|ml|정|캡슐|포|mg)[^)]*)\)\s*$", raw_name)
    if m and ("g" in m.group(2) or "mL" in m.group(2) or "ml" in m.group(2)):
        return m.group(1).strip(), m.group(2).strip()
    return raw_name, ""


def is_export_variant(candidate: dict) -> bool:
    name = str(candidate.get("ITEM_NAME", "")) + str(candidate.get("ITEM_ENG_NAME", ""))
    return ("수출" in name) or ("export" in name.lower())


def pick_best_match(candidates: list, target_name: str, target_content: str = ""):
    """후보 목록 중 제품명(+함량)이 가장 근접한 것을 고른다. 수출용 변형은 후순위로 미룬다."""
    if not candidates:
        return None
    domestic = [c for c in candidates if not is_export_variant(c)]
    pool_all = domestic if domestic else candidates

    target_name_norm = re.sub(r"\s+", "", target_name or "")
    exact = [c for c in pool_all if re.sub(r"\s+", "", c.get("ITEM_NAME", "")) == target_name_norm]
    pool = exact if exact else pool_all

    if target_content:
        content_norm = re.sub(r"\s+", "", str(target_content))
        for c in pool:
            if content_norm and content_norm in re.sub(r"\s+", "", str(c.get("ITEM_NAME", ""))):
                return c
    return pool[0]


TAG_RE = re.compile(r"<[^>]+>")


def strip_html(text: str) -> str:
    if not text:
        return ""
    text = str(text)
    # CDATA 블록은 내부에 '>'가 없어서 태그 제거 정규식이 통째로 지워버리는 문제가 있었음.
    # CDATA 내용을 먼저 그대로 꺼낸 뒤에 남은 XML 태그를 제거한다.
    text = re.sub(r"<!\[CDATA\[(.*?)\]\]>", r"\1", text, flags=re.DOTALL)
    text = TAG_RE.sub(" ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_service_key(raw_key: str, key_mode: str) -> str:
    if key_mode == "인코딩(Encoding) 키를 붙여넣었어요":
        import urllib.parse as up
        try:
            return up.unquote(raw_key)
        except Exception:
            return raw_key
    return raw_key


def call_detail_api(service_key: str, item_seq: str, timeout: int = 15, return_raw: bool = False):
    """ITEM_SEQ 1건에 대해 상세조회 API 호출. drug_api_extractor.py의 call_api와
    동일한 파싱 규칙(response 래핑 유무 모두 지원)을 사용.
    return_raw=True면 (result, err, raw_text, request_url)을 반환 (디버그용)."""
    url = f"{DETAIL_BASE_URL}/{DETAIL_OPERATION}"
    params = {
        "serviceKey": service_key,
        "item_seq": item_seq,
        "numOfRows": 1,
        "pageNo": 1,
        "type": "json",
    }

    def _wrap(result, err, raw_text="", req=None):
        if return_raw:
            return result, err, raw_text, (req.url if req is not None else url)
        return result, err

    try:
        resp = _session.get(url, params=params, timeout=timeout)
    except requests.exceptions.RequestException as e:
        return _wrap(None, f"네트워크 오류: {e}")

    text = resp.text.strip()
    if text.startswith("<"):
        return _wrap(None, f"XML 응답(에러 가능성): {text[:300]}", text, resp)

    try:
        data = resp.json()
    except ValueError:
        return _wrap(None, f"알 수 없는 응답: {text[:300]}", text, resp)

    try:
        root_obj = data["response"] if "response" in data and isinstance(data["response"], dict) else data
        body = root_obj["body"]
        header = root_obj["header"]
        result_code = str(header.get("resultCode", ""))
        if result_code not in ("00", "0"):
            return _wrap(None, f"[{result_code}] {header.get('resultMsg', '')}", text, resp)

        items_raw = body.get("items", "")
        if items_raw in ("", None):
            return _wrap(None, "결과 없음 (해당 품목 상세문서 미등록)", text, resp)
        if isinstance(items_raw, dict) and "item" in items_raw:
            item_val = items_raw["item"]
            item = item_val[0] if isinstance(item_val, list) else item_val
        elif isinstance(items_raw, list):
            item = items_raw[0]
        else:
            return _wrap(None, "예상치 못한 items 구조", text, resp)

        result = {
            "품목기준코드": str(item.get("ITEM_SEQ", item_seq)),
            "효능효과": strip_html(item.get("EE_DOC_DATA", "")),
            "용법용량": strip_html(item.get("UD_DOC_DATA", "")),
            "사용상주의사항_API": strip_html(item.get("NB_DOC_DATA", "")),
        }
        return _wrap(result, None, text, resp)
    except (KeyError, TypeError):
        return _wrap(None, f"예상치 못한 JSON 구조: {str(data)[:400]}", text, resp)


# ============================================================
# 사이드바
# ============================================================
with st.sidebar:
    st.header("🔑 인증 설정")
    service_key_raw = st.text_input("공공데이터포털 서비스키", type="password")
    key_mode = st.radio(
        "붙여넣은 키의 종류",
        ["디코딩(Decoding) 키를 붙여넣었어요", "인코딩(Encoding) 키를 붙여넣었어요"],
        index=0,
    )
    service_key = normalize_service_key(service_key_raw, key_mode)

    st.divider()
    st.header("⚙️ 수집 범위 (청크)")
    st.caption("Streamlit Cloud 타임아웃 방지를 위해 한 번에 500~1000건 권장")
    start_idx = st.number_input("시작 인덱스 (0부터)", min_value=0, value=0, step=100)
    end_idx = st.number_input("종료 인덱스 (미포함)", min_value=1, value=500, step=100)
    sleep_sec = st.slider("요청 간 대기시간(초)", 0.0, 1.0, 0.15, 0.05)
    max_workers = st.slider(
        "동시 처리 수 (병렬)", 1, 10, 5,
        help="품목당 API를 2번(목록조회+상세조회) 호출하는 구조라 순차로는 느립니다. "
             "5~8 정도로 올리면 체감 속도가 크게 빨라집니다. 다만 하루 API 호출 한도는 그대로 소모되니 "
             "한도 자체를 늘리는 게 근본 해결책입니다."
    )

    st.divider()
    if st.button("🗑️ 누적 세션 초기화", use_container_width=True):
        st.session_state.pop("filled_results", None)
        st.success("초기화 완료")

st.title("🧩 효능효과/용법용량 ITEM_SEQ 일괄채우기")
st.caption("Master 엑셀 업로드 → 빈 칸인 품목만 골라 ITEM_SEQ로 하나씩 조회 → 누적 병합")

# ============================================================
# 0) 단일 코드 테스트 (디버그용) - 대량 수집 전에 먼저 이걸로 확인 권장
# ============================================================
with st.expander("🔍 대량 수집 전, 코드 1건 먼저 테스트해보기 (문제 진단용)", expanded=True):
    st.caption(
        "Master 파일의 코드가 식약처 ITEM_SEQ가 아닐 수 있다는 게 확인되어, "
        "이제는 '제품명으로 진짜 ITEM_SEQ 찾기'를 먼저 테스트하는 걸 권장합니다."
    )
    test_item_name = st.text_input("테스트할 제품명 (예: 포크랄시럽(포수클로랄)_(9.5g/95mL) - 포장단위 붙어있어도 자동 정제됨)", key="debug_item_name")
    test_name_run = st.button("제품명으로 진짜 ITEM_SEQ 찾기", disabled=not service_key)
    if not service_key:
        st.warning("사이드바에 서비스키를 먼저 입력해주세요.")

    if test_name_run and test_item_name.strip():
        clean_name, size_hint = clean_product_name(test_item_name.strip())
        candidates, err = resolve_item_seq(service_key, clean_name)
        st.session_state["debug_clean_name"] = clean_name
        st.session_state["debug_size_hint"] = size_hint
        st.session_state["debug_candidates"] = candidates
        st.session_state["debug_list_err"] = err
        # 새로 검색했으니 이전 상세조회 결과는 초기화
        st.session_state.pop("debug_detail_result", None)

    if st.session_state.get("debug_candidates") is not None:
        clean_name = st.session_state.get("debug_clean_name", "")
        size_hint = st.session_state.get("debug_size_hint", "")
        candidates = st.session_state["debug_candidates"]
        err = st.session_state.get("debug_list_err")

        if clean_name and clean_name != test_item_name.strip():
            st.caption(f"🧹 검색용으로 정제된 이름: **{clean_name}** (포장단위 '{size_hint}' 분리함)")

        if err:
            st.error(f"목록조회 실패: {err}")
        elif not candidates:
            st.warning("검색 결과가 없습니다. 제품명 일부만 넣어보세요 (예: 괄호 안 성분명 빼고).")
        else:
            st.success(f"{len(candidates)}건 발견")
            show_cols = [c for c in ["ITEM_SEQ", "ITEM_NAME", "ENTP_NAME", "ITEM_PERMIT_DATE"] if c in candidates[0]]
            st.dataframe(pd.DataFrame(candidates)[show_cols] if show_cols else pd.DataFrame(candidates))

            best = pick_best_match(candidates, clean_name)
            if best:
                real_seq = best.get("ITEM_SEQ")
                st.info(f"가장 유력한 진짜 ITEM_SEQ: **{real_seq}** — 아래 버튼으로 바로 상세조회 테스트")
                detail_test_click = st.button(f"이 ITEM_SEQ({real_seq})로 상세조회 테스트", key="debug_detail_btn")
                if detail_test_click:
                    result, derr, raw_text, req_url = call_detail_api(service_key, real_seq, return_raw=True)
                    st.session_state["debug_detail_result"] = (result, derr, raw_text)

                if "debug_detail_result" in st.session_state:
                    result, derr, raw_text = st.session_state["debug_detail_result"]
                    if result:
                        st.success("성공! 효능효과/용법용량이 반환되었습니다.")
                        st.json(result)
                    else:
                        st.error(f"실패: {derr}")
                    st.write("**원본 응답 전체 (필드명 확인용):**")
                    try:
                        import json as _json
                        st.json(_json.loads(raw_text))
                    except Exception:
                        st.code(raw_text[:3000] if raw_text else "(응답 본문 없음)")

    st.divider()
    st.caption("코드값 자체를 직접 넣어 테스트하고 싶으면 아래를 사용하세요.")
    test_item_seq = st.text_input("테스트할 품목기준코드(ITEM_SEQ)", key="debug_item_seq")
    test_run = st.button("이 코드로 테스트 호출", disabled=not service_key)
    if test_run and test_item_seq.strip():
        t0 = time.time()
        result, err, raw_text, req_url = call_detail_api(
            service_key, test_item_seq.strip(), return_raw=True
        )
        elapsed = time.time() - t0
        st.caption(f"⏱️ 응답 시간: {elapsed:.2f}초")
        st.write("**요청 URL(서비스키 마스킹):**")
        st.code(req_url.split("serviceKey=")[0] + "serviceKey=***" if "serviceKey=" in req_url else req_url)
        if result:
            st.success("성공! 아래 데이터가 반환되었습니다.")
            st.json(result)
        else:
            st.error(f"실패: {err}")
        st.write("**원본 응답(raw, 필드명 확인용):**")
        try:
            import json as _json
            st.json(_json.loads(raw_text))
        except Exception:
            st.code(raw_text[:3000] if raw_text else "(응답 본문 없음)")

st.divider()

# ============================================================
# 1) Master 엑셀 업로드
# ============================================================
master_file = st.file_uploader("Master 엑셀 업로드 (약품명/제품명, 품목기준코드/약품코드, 성분명, 함량, 제형, 효능효과, 용법용량 컬럼 필요)", type=["xlsx"])

if master_file:
    df = pd.read_excel(master_file, sheet_name=0)
    st.write(f"전체 {len(df):,}행 로드 완료")

    name_col = "약품명" if "약품명" in df.columns else ("제품명" if "제품명" in df.columns else None)
    code_col = "약품코드" if "약품코드" in df.columns else ("품목기준코드" if "품목기준코드" in df.columns else None)
    if name_col is None:
        st.error("약품명(또는 제품명) 컬럼을 찾을 수 없습니다. (코드값 불일치가 확인되어 이름으로 재조회하는 방식이 필요합니다)")
        st.stop()

    eff_empty = df["효능효과"].isna() if "효능효과" in df.columns else False
    use_empty = df["용법용량"].isna() if "용법용량" in df.columns else False
    missing_mask = eff_empty | use_empty
    missing_df = df[missing_mask].reset_index(drop=True)
    st.info(
        f"효능효과 또는 용법용량이 비어있는 행: {len(missing_df):,}건 "
        "(용법용량을 못 찾아도 효능효과만이라도 채웁니다)"
    )

    chunk_df = missing_df.iloc[int(start_idx):int(end_idx)]
    st.write(f"이번 청크에서 조회할 품목: {len(chunk_df):,}건 (인덱스 {start_idx}~{end_idx})")
    st.caption(
        "⚠️ 이 방식은 품목당 API를 최대 2번(① 약품명→ITEM_SEQ 검색, ② 상세조회) 호출합니다. "
        "속도가 기존보다 2배 느려지니, 청크 크기를 기존의 절반 정도로 잡는 걸 권장합니다. "
        "약품명 끝에 포장단위(예: '_(9.5g/95mL)')가 붙어있어도 자동으로 떼어내고 검색합니다."
    )

    if "filled_results" not in st.session_state:
        st.session_state["filled_results"] = {}  # 품목기준코드(원본 Master 기준) -> dict
    if "name_seq_cache" not in st.session_state:
        st.session_state["name_seq_cache"] = {}  # 제품명 -> 진짜 ITEM_SEQ (중복 검색 방지)

    run = st.button("🚀 이 범위 수집 시작", type="primary", disabled=not service_key)
    if not service_key:
        st.warning("사이드바에 서비스키를 입력해주세요.")

    if run:
        progress = st.progress(0.0)
        status = st.empty()
        errors = []
        n = len(chunk_df)
        run_start = time.time()

        # 워커 스레드는 st.session_state를 절대 건드리지 않는다 (여러 스레드가 동시에 접근하면
        # Streamlit 세션 상태가 깨져서 앱이 죽는 문제가 있었음). 대신 로컬 dict+락을 쓰고,
        # 세션 상태 갱신은 아래 메인 스레드 루프에서만 한다.
        local_seq_cache = dict(st.session_state["name_seq_cache"])
        cache_lock = Lock()

        def process_row(row_tuple):
            i, row = row_tuple
            orig_code = str(row[code_col]).strip() if code_col else str(i)
            raw_name = str(row[name_col]).strip()
            clean_name, size_hint = clean_product_name(raw_name)
            content_hint = size_hint or (str(row["함량"]).strip() if "함량" in df.columns and pd.notna(row.get("함량")) else "")

            with cache_lock:
                real_seq = local_seq_cache.get(clean_name)

            if real_seq is None:
                candidates, list_err = resolve_item_seq(service_key, clean_name)
                if list_err or not candidates:
                    return ("error", f"{raw_name}: 목록조회 실패/결과없음 ({list_err})", None)
                best = pick_best_match(candidates, clean_name, content_hint)
                real_seq = best.get("ITEM_SEQ") if best else None
                with cache_lock:
                    local_seq_cache[clean_name] = real_seq

            if not real_seq:
                return ("error", f"{raw_name}: 진짜 ITEM_SEQ를 못 찾음", None)

            if sleep_sec > 0:
                time.sleep(sleep_sec)
            result, err = call_detail_api(service_key, real_seq)
            if result:
                result["원본코드"] = orig_code
                return ("ok", orig_code, result)
            return ("error", f"{raw_name}({real_seq}): {err}", None)

        rows = [(i, row) for i, (_, row) in enumerate(chunk_df.iterrows())]
        done_count = 0
        try:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(process_row, rt): rt[0] for rt in rows}
                for future in as_completed(futures):
                    try:
                        outcome = future.result()
                    except Exception as e:
                        errors.append(f"작업 중 예외 발생: {e}")
                        done_count += 1
                        continue
                    if outcome[0] == "ok":
                        _, orig_code, result = outcome
                        st.session_state["filled_results"][orig_code] = result
                    else:
                        errors.append(outcome[1])
                    done_count += 1
                    progress.progress(done_count / max(n, 1))
                    status.text(f"{done_count}/{n} 처리 중... (누적 성공 {len(st.session_state['filled_results']):,}건)")
        finally:
            # 캐시는 정상 종료든 예외든 항상 세션에 반영 (메인 스레드에서만 실행됨)
            st.session_state["name_seq_cache"] = local_seq_cache

        elapsed_total = time.time() - run_start
        rate = n / elapsed_total if elapsed_total > 0 else 0
        st.success(
            f"이번 청크 완료. 누적 수집: {len(st.session_state['filled_results']):,}건 "
            f"(이번 청크 {n:,}건 처리에 {elapsed_total:.0f}초 소요, 초당 {rate:.2f}건 = "
            f"시간당 약 {rate*3600:,.0f}건 페이스)"
        )
        if errors:
            with st.expander(f"⚠️ 실패/미등록 {len(errors)}건"):
                for e in errors[:200]:
                    st.text(e)

    # ============================================================
    # 2) 누적 결과를 원본 Master와 병합
    # ============================================================
    if st.session_state["filled_results"]:
        st.divider()
        st.subheader("📥 누적 결과 병합 다운로드")

        result_df = pd.DataFrame(st.session_state["filled_results"].values())
        merged = df.copy()
        if code_col:
            merged[code_col] = merged[code_col].astype(str).str.strip()
        else:
            merged["_원본코드_임시"] = merged.index.astype(str)
            code_col = "_원본코드_임시"
        result_df["원본코드"] = result_df["원본코드"].astype(str).str.strip()
        result_df = result_df.rename(columns={"품목기준코드": "식약처_진짜ITEM_SEQ"})

        merged = merged.merge(
            result_df, left_on=code_col, right_on="원본코드", how="left", suffixes=("", "_신규")
        )

        if "효능효과_신규" in merged.columns:
            merged["효능효과"] = merged["효능효과"].where(merged["효능효과"].notna() & (merged["효능효과"] != ""), merged["효능효과_신규"])
            merged["용법용량"] = merged["용법용량"].where(merged["용법용량"].notna() & (merged["용법용량"] != ""), merged["용법용량_신규"])
            merged = merged.drop(columns=["효능효과_신규", "용법용량_신규", "품목기준코드"], errors="ignore")

        still_missing = merged["효능효과"].isna().sum() if "효능효과" in merged.columns else 0
        st.write(f"병합 후 여전히 비어있는 행: {still_missing:,}건")

        import io
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="Drug_Master_병합")
        st.download_button(
            "⬇️ 병합된 Master 엑셀 다운로드",
            data=buf.getvalue(),
            file_name="약품_Master_효능효과_채움_병합.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ============================================================
    # 3) 2차 보충: 동일 성분+함량+제형 그룹 상속
    # ============================================================
    st.divider()
    st.subheader("🔁 2차 보충: 동일 성분·함량·제형 그룹 상속")
    st.caption(
        "API로도 채워지지 않는 품목(제네릭이 자체 문서 없이 원개발사 문서를 참조하는 경우 등)에 대해, "
        "같은 성분명+함량+제형 그룹 안에 채워진 값이 있으면 그대로 상속합니다. "
        "'신뢰도' 컬럼에 '성분상속'이라고 표시되므로 원본 API 값과 구분해서 검수할 수 있습니다."
    )
    do_inherit = st.button("성분 그룹 상속 실행 (병합 다운로드 이후 사용 권장)")
    if do_inherit:
        base = merged if "merged" in dir() and isinstance(merged, pd.DataFrame) else df
        base = base.copy()
        if "신뢰도" not in base.columns:
            base["신뢰도"] = ""

        group_cols = [c for c in ["성분명", "함량", "제형"] if c in base.columns]
        if not group_cols:
            st.error("성분명/함량/제형 컬럼이 없어 그룹 상속을 수행할 수 없습니다.")
        else:
            filled_before = base["효능효과"].notna().sum()

            def pick_reference(group: pd.DataFrame):
                ref = group[group["효능효과"].notna() & (group["효능효과"] != "")]
                if ref.empty:
                    return None
                return ref.iloc[0]

            for _, group in base.groupby(group_cols):
                ref_row = pick_reference(group)
                if ref_row is None:
                    continue
                empty_idx = group[group["효능효과"].isna() | (group["효능효과"] == "")].index
                base.loc[empty_idx, "효능효과"] = ref_row["효능효과"]
                base.loc[empty_idx, "용법용량"] = ref_row["용법용량"]
                base.loc[empty_idx, "신뢰도"] = base.loc[empty_idx, "신뢰도"].where(
                    base.loc[empty_idx, "신뢰도"] != "", "성분상속(동일 성분·함량·제형 참조)"
                )

            filled_after = base["효능효과"].notna().sum()
            st.success(f"성분 그룹 상속으로 {filled_after - filled_before:,}건 추가 채움")

            import io
            buf2 = io.BytesIO()
            with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
                base.to_excel(writer, index=False, sheet_name="Drug_Master_최종")
            st.download_button(
                "⬇️ 성분상속 반영 최종 엑셀 다운로드",
                data=buf2.getvalue(),
                file_name="약품_Master_효능효과_최종.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("먼저 Master 엑셀을 업로드해주세요.")

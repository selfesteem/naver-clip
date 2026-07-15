"""Google Sheets 읽기/쓰기 모듈"""
import copy
import json
import os
import time

import gspread
from gspread.utils import rowcol_to_a1
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

TRACKED_SECTIONS = [
    "네이버 클립", "뉴스", "인기글", "이미지", "웹문서 1", "웹문서 2", "플레이스",
]
RESULT_COLS = ["처리완료"] + TRACKED_SECTIONS + ["기타 노출", "오류"]


def _get_client() -> gspread.Client:
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise EnvironmentError("GOOGLE_CREDENTIALS 환경변수가 설정되지 않았습니다.")
    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    return gspread.authorize(creds)


def _open_worksheet(spreadsheet_id: str, gid: int) -> gspread.Worksheet:
    client = _get_client()
    ss = client.open_by_key(spreadsheet_id)
    for ws in ss.worksheets():
        if ws.id == gid:
            return ws
    raise ValueError(f"Sheet GID={gid} 를 찾을 수 없습니다.")


def create_or_get_result_sheet(
    spreadsheet_id: str,
    source_gid: int,
    sheet_name: str,
) -> tuple[int, int]:
    """
    날짜별 결과 시트 탭을 만들거나 이미 있으면 그대로 사용.

    - source_gid  : 키워드 원본 시트 GID
    - sheet_name  : 새 탭 이름 (예: "20260708_오전")

    Returns: (result_sheet_gid, total_keyword_count)
    """
    client = _get_client()
    ss = client.open_by_key(spreadsheet_id)

    # 이미 존재하는지 확인
    existing_ws = None
    for ws in ss.worksheets():
        if ws.title == sheet_name:
            existing_ws = ws
            break

    if existing_ws is not None:
        all_vals = existing_ws.get_all_values()
        total = sum(1 for row in all_vals[1:] if row and row[0].strip())
        if total > 0:
            # 키워드가 있으면 그대로 이어서 처리 (당일 재실행)
            print(f"기존 시트 사용: '{sheet_name}' (키워드 {total}개)")
            return existing_ws.id, total
        # 키워드가 없으면 이전 실행이 중간에 실패한 것 → 아래에서 다시 채움

    # 원본 시트에서 키워드 목록 읽기
    source_ws = None
    for ws in ss.worksheets():
        if ws.id == source_gid:
            source_ws = ws
            break
    if source_ws is None:
        raise ValueError(f"원본 시트 GID={source_gid} 를 찾을 수 없습니다.")

    all_source = source_ws.get_all_values()
    if not all_source:
        raise ValueError("원본 시트가 비어 있습니다.")

    # 키워드 컬럼 찾기
    src_header = all_source[0]
    kw_idx = 0
    for i, h in enumerate(src_header):
        if h in ("키워드", src_header[0]):
            kw_idx = i
            break

    keywords = [
        row[kw_idx].strip()
        for row in all_source[1:]
        if kw_idx < len(row) and row[kw_idx].strip()
    ]

    header_row = ["키워드"] + RESULT_COLS

    if existing_ws is not None:
        # 빈 채로 남은 기존 시트에 데이터만 채우기
        target_ws = existing_ws
        print(f"기존 시트 재초기화: '{sheet_name}'")
    else:
        # 새 시트 생성
        target_ws = ss.add_worksheet(title=sheet_name, rows=len(keywords) + 1, cols=15)
        print(f"새 시트 생성: '{sheet_name}' (키워드 {len(keywords)}개)")

    target_ws.update([header_row], "A1")
    if keywords:
        target_ws.update([[kw] for kw in keywords], "A2")

    return target_ws.id, len(keywords)


class SheetsSession:
    """시트 읽기/쓰기 세션 — 헤더 캐시 + 배치 쓰기 지원."""

    FLUSH_EVERY = 10  # 키워드 N개마다 자동 flush

    def __init__(self, spreadsheet_id: str, gid: int, source_gid: int | None = None):
        self.spreadsheet_id = spreadsheet_id
        self.gid = gid
        self._ws = _open_worksheet(spreadsheet_id, gid)
        self._separate_source = source_gid is not None and source_gid != gid
        self._source_ws = _open_worksheet(spreadsheet_id, source_gid) if self._separate_source else self._ws
        self._refresh_header()
        self._ensure_result_headers()
        self._pending: list[dict] = []
        self._staged_count = 0

    # ── 헤더 관리 ────────────────────────────────────────────────────

    def _refresh_header(self):
        self._header: list[str] = self._ws.row_values(1)
        self._header_map: dict[str, int] = {n: i + 1 for i, n in enumerate(self._header)}

    def _ensure_result_headers(self):
        """결과 컬럼이 없으면 헤더 행 끝에 추가."""
        added = False
        for col in RESULT_COLS:
            if col not in self._header_map:
                self._header.append(col)
                self._header_map[col] = len(self._header)
                added = True
        if added:
            self._ws.update([self._header], "1:1")

    # ── 읽기 ─────────────────────────────────────────────────────────

    def count_keywords(self) -> int:
        """전체 키워드 수 (헤더 제외, 빈 행 제외)."""
        all_values = self._source_ws.get_all_values()
        if len(all_values) < 2:
            return 0
        src_header = all_values[0]
        kw_idx = next((i for i, h in enumerate(src_header) if h == "키워드"), 0)
        return sum(1 for row in all_values[1:] if kw_idx < len(row) and row[kw_idx].strip())

    def read_keywords(self, start: int, count: int | None) -> tuple[list[str], list[int]]:
        """
        지정 범위에서 미완료 키워드만 읽기.
        키워드는 소스 시트에서, 처리완료 여부는 결과 시트에서 확인.

        Returns:
            keywords    : 키워드 목록
            row_indices : 결과 시트 기준 1-based 행 번호
        """
        src_all = self._source_ws.get_all_values()
        if len(src_all) < 2:
            return [], []

        src_header = src_all[0]
        kw_idx = next((i for i, h in enumerate(src_header) if h == "키워드"), 0)
        data_rows = src_all[1:]
        sliced = data_rows[start : (start + count) if count else None]

        # 처리완료 상태는 결과 시트에서 확인
        done_col = self._header_map.get("처리완료")  # 1-based
        if self._separate_source:
            result_all = self._ws.get_all_values()
            result_data = result_all[1:] if len(result_all) > 1 else []
        else:
            result_data = data_rows

        keywords, row_indices = [], []
        for i, row in enumerate(sliced):
            kw = row[kw_idx].strip() if kw_idx < len(row) else ""
            if not kw:
                continue
            done_val = ""
            result_pos = start + i
            if done_col and result_pos < len(result_data):
                r = result_data[result_pos]
                if (done_col - 1) < len(r):
                    done_val = r[done_col - 1].strip()
            if done_val == "Y":
                continue
            keywords.append(kw)
            row_indices.append(start + i + 2)  # +2: 헤더 행 + 1-based 변환

        return keywords, row_indices

    # ── 쓰기 ─────────────────────────────────────────────────────────

    def stage_result(self, row_idx: int, result: dict):
        """결과 한 건을 버퍼에 추가. FLUSH_EVERY에 도달하면 자동 flush."""
        updates: dict[str, str] = {col: "" for col in TRACKED_SECTIONS}
        updates["기타 노출"] = ""
        updates["오류"] = result["error"] or ""

        other_found: list[str] = []
        for sec in result.get("sections", []):
            name: str = sec["name"]
            has: bool = sec["has_target"]
            pos: int | None = sec["position"]
            positions = sec.get("positions") or ([pos] if pos else [])
            cell_val = (",".join(map(str, positions)) if positions else "있음") if has else "X"
            if name in TRACKED_SECTIONS:
                updates[name] = cell_val
            elif has:
                other_found.append(f"{name}:{cell_val}")

        updates["기타 노출"] = ", ".join(other_found)
        updates["처리완료"] = "Y"

        for col_name, value in updates.items():
            if col_name in self._header_map:
                self._pending.append({
                    "range": rowcol_to_a1(row_idx, self._header_map[col_name]),
                    "values": [[value]],
                })

        self._staged_count += 1
        if self._staged_count >= self.FLUSH_EVERY:
            self.flush()

    def flush(self):
        """버퍼를 시트에 일괄 기록."""
        if not self._pending:
            return
        for attempt in range(4):
            try:
                self._ws.batch_update(copy.deepcopy(self._pending), value_input_option="RAW")
                break
            except Exception as e:
                if attempt == 3:
                    raise
                wait = 10 * (2 ** attempt)  # 10s, 20s, 40s
                print(f"\n  Sheets 쓰기 오류 (재시도 {attempt + 1}/3, {wait}초 후): {e}")
                time.sleep(wait)
        self._pending = []
        self._staged_count = 0

    # ── 내부 ─────────────────────────────────────────────────────────

    def _kw_col_idx(self) -> int:
        """키워드 컬럼 0-based 인덱스."""
        for candidate in ["키워드"] + (self._header[:1] if self._header else []):
            if candidate in self._header_map:
                return self._header_map[candidate] - 1
        return 0

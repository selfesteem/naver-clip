#!/usr/bin/env python3
"""
키워드 검색 → 특정 블로그 아이디의 노출 여부 확인 (O/X)

같은 키워드에 묶인 블로그 아이디는 검색 1회로 일괄 체크.

입력 Google Sheets: A열 키워드, B열 블로그 아이디

사용법:
    python blog_exposure.py keywords.xlsx --headless
    python blog_exposure.py --sheets-id ID --sheet-gid GID --headless
    python blog_exposure.py --sheets-id ID --sheet-gid GID --row-indices 2,5,8,10 --headless
"""

import asyncio
import argparse
import os
import sys
import random
from collections import defaultdict
from datetime import date
from pathlib import Path
from urllib.parse import quote

import pandas as pd
from playwright.async_api import async_playwright, Page

from naver_clip import MOBILE_UA

DELAY_MIN = 2.0
DELAY_MAX = 5.0
BATCH_SIZE = 50       # 키워드 기준 배치 크기
BATCH_BREAK_MIN = 15
BATCH_BREAK_MAX = 30
CONTEXT_RESET_EVERY = 200

RESULT_COLS = ["처리완료", "노출여부", "오류"]

USER_AGENTS = [
    MOBILE_UA,
    (
        "Mozilla/5.0 (Linux; Android 14; Pixel 8) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.6367.82 Mobile Safari/537.36"
    ),
    (
        "Mozilla/5.0 (iPhone; CPU iPhone OS 16_6 like Mac OS X) "
        "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.6 Mobile/15E148 Safari/604.1"
    ),
]
VIEWPORTS = [
    {"width": 390, "height": 844},
    {"width": 375, "height": 812},
    {"width": 412, "height": 915},
]

# 블로그 아이디 배열을 받아 DOM 클론 1회로 모두 체크 → {blogId: bool}
_EXPOSURE_JS_MULTI = """
(blogIds) => {
    function esc(s) {
        return s.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
    }

    // a[href] 목록과 body 클론은 한 번만 만들어 재사용
    const hrefs = Array.from(document.querySelectorAll('a[href*="blog.naver.com/"]'))
                       .map(a => a.getAttribute('href') || '');
    const clone = document.body.cloneNode(true);
    clone.querySelectorAll('script, style, noscript, template').forEach(e => e.remove());
    const bodyText = clone.innerHTML;

    const result = {};
    for (const blogId of blogIds) {
        const re = new RegExp('blog\\.naver\\.com/' + esc(blogId) + '([/"\'&?#=\\\\s]|$)', 'i');
        result[blogId] = hrefs.some(h => re.test(h)) || re.test(bodyText);
    }
    return result;
}
"""


def _mask_for_ci(*values: str):
    """GitHub Actions 로그에서 값을 마스킹 (CI 환경에서만 동작)."""
    if not os.environ.get("GITHUB_ACTIONS"):
        return
    for v in values:
        if v:
            print(f"::add-mask::{v}", flush=True)


def _normalize_blog_id(raw: str) -> str:
    """'blog.naver.com/xxx/...' 형태 입력도 아이디만 남김."""
    val = raw.strip()
    if "blog.naver.com/" in val:
        parts = [p for p in val.split("blog.naver.com/")[-1].split("/") if p]
        if parts:
            val = parts[0]
    return val.split("?")[0].strip()


def _build_update(keyword: str, blog_id: str, exposed: bool | None, error: str | None) -> dict[str, str]:
    return {
        "처리완료": "Y",
        "노출여부": "O" if exposed else ("X" if exposed is not None else ""),
        "오류": error or "",
    }


def load_pairs(filepath: str, start: int, count: int | None) -> list[tuple[str, str]]:
    """A열(키워드), B열(아이디) 쌍 목록 읽기."""
    path = Path(filepath)
    if not path.exists():
        sys.exit(f"파일을 찾을 수 없습니다: {filepath}")

    xl = pd.ExcelFile(path)
    df = xl.parse(xl.sheet_names[0], dtype=str)
    if df.empty or df.shape[1] < 2:
        sys.exit("엑셀 파일이 비어있거나 B열(아이디)이 없습니다.")

    kw_col, id_col = df.columns[0], df.columns[1]
    pairs = []
    for kw, bid in zip(df[kw_col], df[id_col]):
        kw = str(kw).strip() if pd.notna(kw) else ""
        bid = _normalize_blog_id(str(bid)) if pd.notna(bid) else ""
        if kw and bid:
            pairs.append((kw, bid))
    return pairs[start : (start + count) if count else None]


def _group_by_keyword(pairs: list[tuple[str, str]]) -> dict[str, list[str]]:
    """(keyword, blog_id) 쌍을 keyword → [blog_id, ...] 로 묶음. 순서 유지."""
    groups: dict[str, list[str]] = defaultdict(list)
    for kw, bid in pairs:
        groups[kw].append(bid)
    return dict(groups)


def load_or_create_output(output_path: Path, pairs: list[tuple[str, str]]) -> pd.DataFrame:
    if output_path.exists():
        df = pd.read_excel(output_path, dtype=str)
        for c in RESULT_COLS:
            if c not in df.columns:
                df[c] = ""
        return df

    df = pd.DataFrame(pairs, columns=["키워드", "아이디"])
    for c in RESULT_COLS:
        df[c] = ""
    return df


def save_dataframe(df: pd.DataFrame, filepath: Path):
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="노출여부")
        ws = writer.sheets["노출여부"]
        for col_cells in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col_cells), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 40)


def mark_results(df: pd.DataFrame, keyword: str, exposures: dict[str, bool | None], error: str | None):
    """키워드에 속한 모든 아이디의 결과를 df에 반영."""
    for bid, exposed in exposures.items():
        mask = (
            (df["키워드"].astype(str).str.strip() == keyword) &
            (df["아이디"].astype(str).str.strip() == bid)
        )
        for i in df[mask].index[:1]:
            for col, val in _build_update(keyword, bid, exposed, error).items():
                if col in df.columns:
                    df.at[i, col] = val


def is_done(val) -> bool:
    return pd.notna(val) and str(val).strip() == "Y"


async def search_keyword_exposure(page: Page, keyword: str, blog_ids: list[str]) -> dict:
    """키워드 검색 1회 → 여러 블로그 아이디 노출 여부를 한꺼번에 반환."""
    result = {"keyword": keyword, "exposures": {bid: None for bid in blog_ids}, "error": None}
    try:
        url = f"https://m.search.naver.com/search.naver?query={quote(keyword)}"
        await page.goto(url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(2_000)
        result["exposures"] = await page.evaluate(_EXPOSURE_JS_MULTI, blog_ids)
    except Exception as e:
        msg = str(e)
        for secret in (keyword, quote(keyword), *blog_ids):
            if secret and secret in msg:
                msg = msg.replace(secret, "***")
        result["error"] = msg
    return result


async def make_context(browser):
    ctx = await browser.new_context(
        user_agent=random.choice(USER_AGENTS),
        viewport=random.choice(VIEWPORTS),
        locale="ko-KR",
    )
    await ctx.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return ctx


async def run(input_file: str, headless: bool, start: int,
              count: int | None, output_dir: str | None):

    pairs = load_pairs(input_file, start, count)
    if not pairs:
        sys.exit(f"범위(start={start})에 해당하는 키워드/아이디 쌍이 없습니다.")

    for kw, bid in pairs:
        _mask_for_ci(kw, bid)

    today = date.today().strftime("%Y%m%d")
    actual_pairs = len(pairs)
    out_dir = Path(output_dir) if output_dir else Path(input_file).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / f"blog_exposure_{today}_{start}_{actual_pairs}.xlsx"

    df = load_or_create_output(output_path, pairs)

    done_keys = {
        (str(k).strip(), str(b).strip())
        for k, b in zip(df.loc[df["처리완료"].apply(is_done), "키워드"],
                        df.loc[df["처리완료"].apply(is_done), "아이디"])
    }
    pending_pairs = [(kw, bid) for kw, bid in pairs if (kw, bid) not in done_keys]
    pending_groups = _group_by_keyword(pending_pairs)
    already_done = actual_pairs - len(pending_pairs)

    if already_done:
        print(f"이미 완료: {already_done}개 → 남은 {len(pending_pairs)}개 처리")
    if not pending_groups:
        print("모든 키워드가 처리되었습니다.")
        return

    unique_kw = len(pending_groups)
    print(f"총 {actual_pairs}쌍 | 처리 예정: {len(pending_pairs)}쌍 ({unique_kw}개 키워드)")
    eta = unique_kw * ((DELAY_MIN + DELAY_MAX) / 2)
    print(f"예상 소요 시간: 약 {int(eta // 3600)}시간 {int((eta % 3600) // 60)}분\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        context = await make_context(browser)
        page = await context.new_page()

        for kw_idx, (kw, blog_ids) in enumerate(pending_groups.items()):
            overall = already_done + sum(len(v) for v in list(pending_groups.values())[:kw_idx]) + 1
            print(f"[{kw_idx + 1:>5}/{unique_kw}] ({len(blog_ids)}개) ...", end=" ", flush=True)

            result = await search_keyword_exposure(page, kw, blog_ids)
            mark_results(df, kw, result["exposures"], result["error"])
            save_dataframe(df, output_path)

            if result["error"]:
                print(f"오류: {result['error']}")
            else:
                summary = " ".join(
                    f"{'O' if v else 'X'}" for v in result["exposures"].values()
                )
                print(summary)

            if kw_idx == unique_kw - 1:
                break

            if (kw_idx + 1) % BATCH_SIZE == 0:
                pause = random.uniform(BATCH_BREAK_MIN, BATCH_BREAK_MAX)
                print(f"\n  {BATCH_SIZE}개 완료 — {pause:.0f}초 휴식...\n")
                await asyncio.sleep(pause)
                if (kw_idx + 1) % CONTEXT_RESET_EVERY == 0:
                    await context.close()
                    context = await make_context(browser)
                    page = await context.new_page()
            else:
                await asyncio.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

        await browser.close()

    done_df = df[df["처리완료"] == "Y"]
    exposed_cnt = (done_df["노출여부"] == "O").sum() if "노출여부" in done_df.columns else 0
    print(f"\n결과 저장: {output_path}")
    print(f"\n--- 요약 ({len(done_df)}/{actual_pairs}) ---")
    print(f"  노출: {exposed_cnt}개 / 미노출: {len(done_df) - exposed_cnt}개")


# ── Google Sheets 지원 ────────────────────────────────────────────────────────

from sheets_io import _api_call as _sh, _get_client as _sh_client  # noqa: E402


class BlogExposureSheetsSession:
    """Google Sheets 블로그 노출 여부 데이터 읽기/쓰기 세션."""

    FLUSH_EVERY = 10

    def __init__(self, spreadsheet_id: str, gid: int, source_gid: int | None = None):
        from gspread.utils import rowcol_to_a1
        self._rowcol_to_a1 = rowcol_to_a1
        self._separate_source = source_gid is not None and source_gid != gid

        ss = _sh_client().open_by_key(spreadsheet_id)
        ws_map = {ws.id: ws for ws in _sh(ss.worksheets)}

        if gid not in ws_map:
            raise ValueError(f"Sheet GID={gid} 를 찾을 수 없습니다.")
        self._ws = ws_map[gid]

        if self._separate_source:
            if source_gid not in ws_map:
                raise ValueError(f"Sheet GID={source_gid} 를 찾을 수 없습니다.")
            self._source_ws = ws_map[source_gid]
        else:
            self._source_ws = self._ws

        self._header: list[str] = _sh(self._ws.row_values, 1)
        self._header_map: dict[str, int] = {n: i + 1 for i, n in enumerate(self._header)}
        self._pending: list[dict] = []
        self._staged_count = 0

    def read_pairs(self, start: int, count: int | None) -> tuple[list[tuple[str, str]], list[int]]:
        """지정 범위에서 미완료 (키워드, 아이디) 쌍만 읽기."""
        src_all = _sh(self._source_ws.get_all_values)
        if len(src_all) < 2:
            return [], []
        src_header = src_all[0]
        kw_idx = next((i for i, h in enumerate(src_header) if h == "키워드"), 0)
        id_idx = next((i for i, h in enumerate(src_header) if h == "아이디"), 1)
        data_rows = src_all[1:]
        sliced = data_rows[start : (start + count) if count else None]

        done_col = self._header_map.get("처리완료")
        if self._separate_source:
            result_all = _sh(self._ws.get_all_values)
            result_data = result_all[1:] if len(result_all) > 1 else []
        else:
            result_data = data_rows

        pairs, row_indices = [], []
        for i, row in enumerate(sliced):
            kw = row[kw_idx].strip() if kw_idx < len(row) else ""
            bid = _normalize_blog_id(row[id_idx]) if id_idx < len(row) else ""
            if not kw or not bid:
                continue
            done_val = ""
            result_pos = start + i
            if done_col and result_pos < len(result_data):
                r = result_data[result_pos]
                if (done_col - 1) < len(r):
                    done_val = r[done_col - 1].strip()
            if done_val == "Y":
                continue
            pairs.append((kw, bid))
            row_indices.append(start + i + 2)

        return pairs, row_indices

    def read_pairs_by_indices(self, row_indices: list[int]) -> tuple[list[tuple[str, str]], list[int]]:
        """지정된 시트 행 번호(1-based)의 미완료 쌍만 읽기."""
        all_vals = _sh(self._ws.get_all_values)
        if len(all_vals) < 2:
            return [], []
        header = all_vals[0]
        kw_idx = next((i for i, h in enumerate(header) if h == "키워드"), 0)
        id_idx = next((i for i, h in enumerate(header) if h == "아이디"), 1)
        done_col = self._header_map.get("처리완료")

        pairs, out_indices = [], []
        for row_num in row_indices:
            sheet_idx = row_num - 1  # 0-based into all_vals
            if sheet_idx <= 0 or sheet_idx >= len(all_vals):
                continue
            row = all_vals[sheet_idx]
            kw = row[kw_idx].strip() if kw_idx < len(row) else ""
            bid = _normalize_blog_id(row[id_idx]) if id_idx < len(row) else ""
            if not kw or not bid:
                continue
            done_val = ""
            if done_col and (done_col - 1) < len(row):
                done_val = row[done_col - 1].strip()
            if done_val == "Y":
                continue
            pairs.append((kw, bid))
            out_indices.append(row_num)

        return pairs, out_indices

    def stage_result(self, row_idx: int, keyword: str, blog_id: str,
                     exposed: bool | None, error: str | None):
        """결과 한 건을 버퍼에 추가. FLUSH_EVERY에 도달하면 자동 flush."""
        for col_name, value in _build_update(keyword, blog_id, exposed, error).items():
            if col_name in self._header_map:
                self._pending.append({
                    "range": self._rowcol_to_a1(row_idx, self._header_map[col_name]),
                    "values": [[value]],
                })

        self._staged_count += 1
        if self._staged_count >= self.FLUSH_EVERY:
            self.flush()

    def flush(self):
        if not self._pending:
            return
        from collections import defaultdict as dd
        from gspread.utils import a1_to_rowcol
        rows: dict = dd(dict)
        for item in self._pending:
            row, col = a1_to_rowcol(item["range"])
            rows[row][col] = item["values"][0][0]
        for row_num in sorted(rows):
            col_vals = rows[row_num]
            min_col, max_col = min(col_vals), max(col_vals)
            values = [[col_vals.get(c, "") for c in range(min_col, max_col + 1)]]
            start_cell = self._rowcol_to_a1(row_num, min_col)
            end_cell = self._rowcol_to_a1(row_num, max_col)
            _sh(self._ws.update, f"{start_cell}:{end_cell}", values, value_input_option="RAW")
        self._pending = []
        self._staged_count = 0


async def run_sheets(spreadsheet_id: str, gid: int, headless: bool,
                     start: int, count: int | None,
                     source_gid: int | None = None,
                     row_indices: list[int] | None = None):
    """Google Sheets 모드: 소스 시트에서 읽고 결과 시트에 씀."""
    session = BlogExposureSheetsSession(spreadsheet_id, gid, source_gid=source_gid or None)

    if row_indices is not None:
        pairs, pair_row_indices = session.read_pairs_by_indices(row_indices)
    else:
        pairs, pair_row_indices = session.read_pairs(start, count)

    for kw, bid in pairs:
        _mask_for_ci(kw, bid)

    if not pairs:
        print("처리할 키워드/아이디 쌍이 없습니다.")
        return

    # 같은 키워드는 한 번만 검색, row_index 정보도 함께 유지
    keyword_groups: dict[str, list[tuple[str, int]]] = defaultdict(list)
    for (kw, bid), row_idx in zip(pairs, pair_row_indices):
        keyword_groups[kw].append((bid, row_idx))

    unique_kw = len(keyword_groups)
    total_pairs = len(pairs)
    print(f"Google Sheets 모드 | 처리 예정: {total_pairs}쌍 ({unique_kw}개 키워드)")
    eta = unique_kw * ((DELAY_MIN + DELAY_MAX) / 2)
    print(f"예상 소요 시간: 약 {int(eta // 3600)}시간 {int((eta % 3600) // 60)}분\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        context = await make_context(browser)
        page = await context.new_page()

        for kw_idx, (kw, bid_rows) in enumerate(keyword_groups.items()):
            blog_ids = [bid for bid, _ in bid_rows]
            print(f"[{kw_idx + 1:>5}/{unique_kw}] ({len(blog_ids)}개) ...", end=" ", flush=True)

            result = await search_keyword_exposure(page, kw, blog_ids)

            for bid, row_idx in bid_rows:
                exposed = result["exposures"].get(bid)
                session.stage_result(row_idx, kw, bid, exposed, result["error"])

            if result["error"]:
                print(f"오류: {result['error']}")
            else:
                summary = " ".join("O" if result["exposures"].get(b) else "X" for b in blog_ids)
                print(summary)

            if kw_idx == unique_kw - 1:
                break

            if (kw_idx + 1) % BATCH_SIZE == 0:
                session.flush()
                pause = random.uniform(BATCH_BREAK_MIN, BATCH_BREAK_MAX)
                print(f"\n  {BATCH_SIZE}개 완료 — {pause:.0f}초 휴식...\n")
                await asyncio.sleep(pause)
                if (kw_idx + 1) % CONTEXT_RESET_EVERY == 0:
                    await context.close()
                    context = await make_context(browser)
                    page = await context.new_page()
            else:
                await asyncio.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

        await browser.close()

    session.flush()
    print(f"\n완료: {total_pairs}쌍 처리 → 구글 시트에 저장됨")


def main():
    parser = argparse.ArgumentParser(description="키워드 검색 → 블로그 아이디 노출 여부 확인 (O/X)")
    parser.add_argument("input_file", nargs="?", help="키워드(A열)/아이디(B열) Excel 파일 경로")
    parser.add_argument("--start", type=int, default=0, help="시작 행 번호 (0-based, 기본: 0)")
    parser.add_argument("--count", type=int, default=None, help="처리할 행 수 (기본: 전체)")
    parser.add_argument("--output-dir", help="결과 파일 저장 폴더 (기본: 입력 파일과 동일)")
    parser.add_argument("--headless", action="store_true", help="브라우저 숨김 모드")
    # Google Sheets 모드
    parser.add_argument("--sheets-id", help="Google Spreadsheet ID")
    parser.add_argument("--sheet-gid", type=int, default=0, help="결과 시트 GID")
    parser.add_argument("--source-gid", type=int, default=None, help="키워드 소스 시트 GID")
    parser.add_argument("--row-indices", help="처리할 시트 행 번호 (쉼표 구분, 1-based)")
    args = parser.parse_args()

    if args.sheets_id:
        row_indices = (
            [int(r) for r in args.row_indices.split(",") if r.strip()]
            if args.row_indices else None
        )
        asyncio.run(run_sheets(
            args.sheets_id, args.sheet_gid, args.headless,
            args.start, args.count,
            source_gid=args.source_gid or None,
            row_indices=row_indices,
        ))
    elif args.input_file:
        asyncio.run(run(args.input_file, args.headless, args.start, args.count, args.output_dir))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

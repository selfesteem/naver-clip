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
from datetime import date, datetime, timedelta, timezone
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
        const re = new RegExp(`blog\\\\.naver\\\\.com/${esc(blogId)}([/"'&?#=\\\\s]|$)`, 'i');
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

KST = timezone(timedelta(hours=9))


def _parse_date(s: str) -> date | None:
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            pass
    return None


class BlogExposureSheetsSession:
    """Google Sheets 블로그 노출 여부 읽기/쓰기 세션.

    시트 구조:
      A(1)=발행일자  B(2)=메인키워드  D(4)=서브키워드  N(14)=URL
      F(6)=당일_메인  G(7)=당일_서브
      H(8)=다음날_메인  I(9)=다음날_서브
      J(10)=1주일뒤_메인  K(11)=1주일뒤_서브
    노출 시 'O' 기록, 미노출은 빈칸 유지.
    """

    FLUSH_EVERY = 10

    COL_DATE = 1
    COL_MAIN_KW = 2
    COL_SUB_KW = 4
    COL_URL = 14

    # offset(일) → (메인결과열, 서브결과열) 1-based
    DATE_OFFSET_COLS = {
        0: (6, 7),    # F, G  발행당일
        1: (8, 9),    # H, I  다음날
        7: (10, 11),  # J, K  1주일뒤
    }

    def __init__(self, spreadsheet_id: str, gid: int):
        from gspread.utils import rowcol_to_a1
        self._rowcol_to_a1 = rowcol_to_a1

        ss = _sh_client().open_by_key(spreadsheet_id)
        ws_map = {ws.id: ws for ws in _sh(ss.worksheets)}
        if gid not in ws_map:
            raise ValueError(f"Sheet GID={gid} 를 찾을 수 없습니다.")
        self._ws = ws_map[gid]
        self._pending: list[dict] = []
        self._staged_count = 0

    def read_pending_rows(self, row_indices: list[int] | None = None) -> list[dict]:
        """오늘 날짜가 발행일+{0,1,7}인 미완료 행을 반환.

        각 항목: {row_idx, main_kw, sub_kw, blog_id, main_col, sub_col, main_done, sub_done}
        """
        today = datetime.now(KST).date()
        all_vals = _sh(self._ws.get_all_values)
        row_index_set = set(row_indices) if row_indices else None
        rows = []

        for i, row in enumerate(all_vals[1:]):
            row_num = i + 2
            if row_index_set and row_num not in row_index_set:
                continue

            date_str = row[self.COL_DATE - 1].strip() if len(row) >= self.COL_DATE else ""
            pub_date = _parse_date(date_str)
            if pub_date is None:
                continue

            url = row[self.COL_URL - 1].strip() if len(row) >= self.COL_URL else ""
            blog_id = _normalize_blog_id(url) if url else ""
            if not blog_id:
                continue

            for offset, (main_col, sub_col) in self.DATE_OFFSET_COLS.items():
                if today != pub_date + timedelta(days=offset):
                    continue

                main_kw = row[self.COL_MAIN_KW - 1].strip() if len(row) >= self.COL_MAIN_KW else ""
                sub_kw = row[self.COL_SUB_KW - 1].strip() if len(row) >= self.COL_SUB_KW else ""
                cur_main = row[main_col - 1].strip() if len(row) >= main_col else ""
                cur_sub = row[sub_col - 1].strip() if len(row) >= sub_col else ""

                rows.append({
                    "row_idx": row_num,
                    "main_kw": main_kw,
                    "sub_kw": sub_kw,
                    "blog_id": blog_id,
                    "main_col": main_col,
                    "sub_col": sub_col,
                    "main_done": bool(cur_main),
                    "sub_done": bool(cur_sub),
                })
                break

        return rows

    def stage_result(self, row_idx: int, col: int):
        """노출 확인 시 'O' 기록 (미노출은 호출하지 않음)."""
        self._pending.append({
            "range": self._rowcol_to_a1(row_idx, col),
            "values": [["O"]],
        })
        self._staged_count += 1
        if self._staged_count >= self.FLUSH_EVERY:
            self.flush()

    def flush(self):
        if not self._pending:
            return
        from gspread.utils import a1_to_rowcol
        rows: dict = defaultdict(dict)
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
                     row_indices: list[int] | None = None):
    """Google Sheets 모드: 날짜 기준 메인/서브 키워드별 노출 여부 확인 후 기록."""
    session = BlogExposureSheetsSession(spreadsheet_id, gid)
    rows = session.read_pending_rows(row_indices)

    if not rows:
        print("처리할 행이 없습니다.")
        return

    # 같은 키워드는 한 번만 검색: keyword → [(row_idx, blog_id, col), ...]
    search_groups: dict[str, list[tuple[int, str, int]]] = defaultdict(list)
    for row in rows:
        if row["main_kw"] and not row["main_done"]:
            search_groups[row["main_kw"]].append((row["row_idx"], row["blog_id"], row["main_col"]))
        if row["sub_kw"] and not row["sub_done"]:
            search_groups[row["sub_kw"]].append((row["row_idx"], row["blog_id"], row["sub_col"]))

    for row in rows:
        _mask_for_ci(row["main_kw"], row["sub_kw"], row["blog_id"])

    unique_kw = len(search_groups)
    if not unique_kw:
        print("모든 행이 이미 처리되었습니다.")
        return

    print(f"처리 예정: {len(rows)}행 → {unique_kw}회 검색")
    eta = unique_kw * ((DELAY_MIN + DELAY_MAX) / 2)
    print(f"예상 소요 시간: 약 {int(eta // 3600)}시간 {int((eta % 3600) // 60)}분\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        context = await make_context(browser)
        page = await context.new_page()

        for kw_idx, (kw, targets) in enumerate(search_groups.items()):
            blog_ids = list({t[1] for t in targets})
            print(f"[{kw_idx + 1:>5}/{unique_kw}] '{kw}' ({len(blog_ids)}개) ...", end=" ", flush=True)

            result = await search_keyword_exposure(page, kw, blog_ids)

            for row_idx, blog_id, col in targets:
                exposed = result["exposures"].get(blog_id)
                if exposed and not result["error"]:
                    session.stage_result(row_idx, col)

            if result["error"]:
                print(f"오류: {result['error']}")
            else:
                summary = " ".join("O" if result["exposures"].get(t[1]) else "-" for t in targets)
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
    print(f"\n완료: {len(rows)}행 처리됨")


def main():
    parser = argparse.ArgumentParser(description="키워드 검색 → 블로그 URL 노출 여부 확인 (O/빈칸)")
    parser.add_argument("input_file", nargs="?", help="키워드/아이디 Excel 파일 경로 (로컬 모드)")
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--count", type=int, default=None)
    parser.add_argument("--output-dir")
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--sheets-id")
    parser.add_argument("--sheet-gid", type=int, default=0)
    parser.add_argument("--row-indices", help="처리할 시트 행 번호 (쉼표 구분, 1-based)")
    args = parser.parse_args()

    if args.sheets_id:
        row_indices = (
            [int(r) for r in args.row_indices.split(",") if r.strip()]
            if args.row_indices else None
        )
        asyncio.run(run_sheets(
            args.sheets_id, args.sheet_gid, args.headless,
            row_indices=row_indices,
        ))
    elif args.input_file:
        asyncio.run(run(args.input_file, args.headless, args.start, args.count, args.output_dir))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

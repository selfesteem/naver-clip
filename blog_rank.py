#!/usr/bin/env python3
"""
키워드 검색 → 특정 블로그 아이디의 섹션별 순위 확인

입력 Excel: A열 키워드, B열 블로그 아이디

사용법:
    python blog_rank.py keywords.xlsx --headless
    python blog_rank.py keywords.xlsx --start 0 --count 100 --headless
"""

import asyncio
import argparse
import sys
import random
from datetime import date
from pathlib import Path
from urllib.parse import quote

import pandas as pd
from playwright.async_api import async_playwright, Page

from naver_clip import MOBILE_UA, MOBILE_VIEWPORT, normalize_section

DELAY_MIN = 2.0
DELAY_MAX = 5.0
BATCH_SIZE = 50
BATCH_BREAK_MIN = 15
BATCH_BREAK_MAX = 30
CONTEXT_RESET_EVERY = 200

TRACKED_SECTIONS = [
    "블로그", "인플루언서", "인기글", "웹문서 1", "웹문서 2",
]
RESULT_COLS = ["처리완료"] + TRACKED_SECTIONS + ["기타 노출", "오류"]

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

# 각 섹션 아이템에서 blog.naver.com/{아이디}/ 링크를 찾아 순위 반환하는 JS
_BLOG_RANK_JS = """
(blogId) => {
    function esc(s) {
        return s.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
    }
    const re = new RegExp('blog\\\\.naver\\\\.com/' + esc(blogId) + '([/"\\'&?#=\\s]|$)');

    // script/style 제거 후 innerHTML 검사 (JSON 오탐 방지) + a[href] 속성 직접 검사
    // 아이템 자체가 <a>인 경우(메이트 프로필 카드 등) 자기 href도 검사
    function hasBlog(el) {
        if (el.tagName === 'A' && re.test(el.getAttribute('href') || '')) return true;
        const clone = el.cloneNode(true);
        clone.querySelectorAll('script, style').forEach(e => e.remove());
        if (re.test(clone.innerHTML)) return true;
        return Array.from(el.querySelectorAll('a[href]')).some(a =>
            re.test(a.getAttribute('href') || '')
        );
    }

    const results = [];

    // 1. h2 레이블이 있는 named 섹션 (중첩 BX 제외 — AI 브리핑 등 이중 카운트 방지)
    const namedEls = Array.from(
        document.querySelectorAll('[class*="api_subject_bx"], .sc_new, [class*="sc_new"]')
    ).filter(el => {
        if (el.closest('[class*="api_subject_bx"], .sc_new, [class*="sc_new"]') !== el) return false;
        const h2 = el.querySelector('h2');
        return h2 && h2.textContent.trim();
    });

    function d2children(root) {
        const out = [];
        for (const c of root.children)
            for (const cc of c.children) out.push(cc);
        return out;
    }

    // UL/OL은 LI가 아이템, 단일 자식 체인은 손자로 확장 (lst_total 블로그 목록 등)
    function unwrapLists(items) {
        const out = [];
        for (const c of items) {
            if (c.tagName === 'UL' || c.tagName === 'OL') {
                Array.from(c.children).forEach(li => out.push(li));
            } else {
                out.push(c);
            }
        }
        return out;
    }

    // 스크립트/빈 껍데기 항목 제외 (fsolid_list는 아이템 DIV와 <script>가 번갈아 있음)
    function hasVisibleText(el) {
        if (['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE'].includes(el.tagName)) return false;
        const clone = el.cloneNode(true);
        clone.querySelectorAll('script, style').forEach(e => e.remove());
        return (clone.textContent || '').trim().length > 0;
    }

    function collectItems(el) {
        let items = unwrapLists(Array.from(el.children));
        for (let depth = 0; depth < 3 && items.length <= 2; depth++) {
            let next = [];
            for (const c of items) next = next.concat(Array.from(c.children));
            next = unwrapLists(next);
            if (next.length <= items.length) break;
            items = next;
        }
        return items.filter(hasVisibleText);
    }

    for (const sec of namedEls) {
        // 블로그 링크가 하나도 없는 섹션(클립/뉴스/이미지/쇼핑 등)은 스킵
        if (!sec.querySelector('a[href*="blog.naver.com/"]')) continue;

        const h2 = sec.querySelector('h2');
        const name = h2.textContent.trim().replace(/\\s+/g, ' ');

        const d2 = d2children(sec);
        if (d2.length < 2) {
            results.push({ name, has_target: hasBlog(sec), position: null, positions: [], total: 0 });
            continue;
        }

        const items = collectItems(d2[1]);

        const positions = [];
        for (let i = 0; i < items.length; i++) {
            if (hasBlog(items[i])) positions.push(i + 1);
        }

        results.push({
            name,
            has_target: positions.length > 0 || (items.length === 0 && hasBlog(sec)),
            position: positions[0] || null,
            positions,
            total: items.length,
        });
    }

    // 2. 웹문서 섹션 — spw_fsolid 방식과 fds-web-list-root 방식 둘 다 처리, DOM 순서 유지
    const webDocEntries = [];

    document.querySelectorAll('.spw_fsolid').forEach(sec => {
        // 블로그 링크가 하나도 없으면 스킵
        if (!sec.querySelector('a[href*="blog.naver.com/"]')) return;
        const list = sec.querySelector('.fsolid_list');
        const items = collectItems(list || sec);
        if (items.length > 0) webDocEntries.push({ el: sec, items });
    });

    document.querySelectorAll('[class*="fds-web-list-root"]').forEach(root => {
        // 블로그 링크가 하나도 없으면 스킵
        if (!root.querySelector('a[href*="blog.naver.com/"]')) return;
        const parentBx = root.closest('[class*="api_subject_bx"]');
        if (!parentBx) return;
        const h2 = parentBx.querySelector('h2');
        if (h2 && h2.textContent.trim()) return;
        const items = Array.from(root.children).filter(el =>
            el.className && el.className.includes('fds-web-doc-') && el.textContent.trim().length > 0
        );
        if (items.length >= 2) webDocEntries.push({ el: root, items });
    });

    webDocEntries.sort((a, b) =>
        a.el.compareDocumentPosition(b.el) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1
    );

    webDocEntries.forEach(({ items }, idx) => {
        const positions = [];
        for (let i = 0; i < items.length; i++) {
            if (hasBlog(items[i])) positions.push(i + 1);
        }
        results.push({
            name: `웹문서 ${idx + 1}`,
            has_target: positions.length > 0,
            position: positions[0] || null,
            positions,
            total: items.length,
        });
    });

    return results;
}
"""


def _normalize_blog_id(raw: str) -> str:
    """'blog.naver.com/xxx/...' 형태 입력도 아이디만 남김."""
    val = raw.strip()
    if "blog.naver.com/" in val:
        parts = [p for p in val.split("blog.naver.com/")[-1].split("/") if p]
        if parts:
            val = parts[0]
    return val.split("?")[0].strip()


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


def load_or_create_output(output_path: Path, pairs: list[tuple[str, str]]) -> pd.DataFrame:
    """결과 파일 로드(재시작 시) 또는 신규 생성."""
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
        df.to_excel(writer, index=False, sheet_name="블로그순위")
        ws = writer.sheets["블로그순위"]
        for col_cells in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col_cells), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 40)


def _merge_sections(sections: list[dict]) -> dict[str, dict]:
    """같은 이름으로 여러 개 열거된 섹션(AI 브리핑 등)의 positions 병합."""
    merged: dict[str, dict] = {}
    for sec in sections:
        name = sec["name"]
        if name not in merged:
            merged[name] = {
                "has_target": bool(sec.get("has_target")),
                "positions": list(sec.get("positions") or []),
            }
        else:
            m = merged[name]
            m["has_target"] = m["has_target"] or bool(sec.get("has_target"))
            for p in sec.get("positions") or []:
                if p not in m["positions"]:
                    m["positions"].append(p)
    return merged


def _build_updates(result: dict) -> dict[str, str]:
    """검색 결과 → 컬럼별 셀 값 (엑셀/시트 공용)."""
    updates: dict[str, str] = {col: "" for col in TRACKED_SECTIONS}
    updates["기타 노출"] = ""
    updates["오류"] = result["error"] or ""

    other_found: list[str] = []
    for name, sec in _merge_sections(result.get("sections", [])).items():
        has: bool = sec["has_target"]
        positions = sec["positions"]
        cell_val = (",".join(map(str, positions)) if positions else "있음") if has else "X"

        if name in TRACKED_SECTIONS:
            updates[name] = cell_val
        elif has:
            other_found.append(f"{name}:{cell_val}")

    updates["기타 노출"] = ", ".join(other_found)
    updates["처리완료"] = "Y"
    return updates


def mark_result(df: pd.DataFrame, result: dict):
    kw, bid = result["keyword"], result["blog_id"]
    mask = (df["키워드"].astype(str).str.strip() == kw) & (df["아이디"].astype(str).str.strip() == bid)
    idxs = df[mask].index
    if len(idxs) == 0:
        return
    i = idxs[0]

    for col in TRACKED_SECTIONS:
        df.at[i, col] = ""
    df.at[i, "기타 노출"] = ""
    df.at[i, "오류"] = result["error"] or ""

    for col, val in _build_updates(result).items():
        if col in df.columns:
            df.at[i, col] = val


def is_done(val) -> bool:
    return pd.notna(val) and str(val).strip() == "Y"


def _brief(sections: list[dict]) -> str:
    found = []
    for name, s in _merge_sections(sections).items():
        if s["has_target"]:
            if s["positions"]:
                found.append(f"{name} {','.join(map(str, s['positions']))}위")
            else:
                found.append(f"{name} 있음")
    return " / ".join(found) if found else "미노출"


async def search_blog_rank(page: Page, keyword: str, blog_id: str) -> dict:
    """모바일 네이버에서 keyword 검색 → blog_id의 섹션별 순위 반환."""
    result = {"keyword": keyword, "blog_id": blog_id, "sections": [], "error": None}
    try:
        url = f"https://m.search.naver.com/search.naver?query={quote(keyword)}"
        await page.goto(url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(2_000)

        raw = await page.evaluate(_BLOG_RANK_JS, blog_id)
        for sec in raw:
            sec["name"] = normalize_section(sec["name"])
        result["sections"] = raw

    except Exception as e:
        msg = str(e)
        for secret in (keyword, blog_id, quote(keyword)):
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

    today = date.today().strftime("%Y%m%d")
    actual_count = len(pairs)
    out_dir = Path(output_dir) if output_dir else Path(input_file).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / f"blog_rank_{today}_{start}_{actual_count}.xlsx"

    df = load_or_create_output(output_path, pairs)

    done_mask = df["처리완료"].apply(is_done)
    already_done = int(done_mask.sum())
    done_keys = {
        (str(k).strip(), str(b).strip())
        for k, b in zip(df.loc[done_mask, "키워드"], df.loc[done_mask, "아이디"])
    }
    pending = [(kw, bid) for kw, bid in pairs if (kw, bid) not in done_keys]

    print(f"출력 파일: {output_path}")
    if already_done:
        print(f"이미 완료: {already_done}개 → 남은 {len(pending)}개 처리")
    if not pending:
        print("모든 키워드가 처리되었습니다.")
        return

    print(f"총 {actual_count}개 | 처리 예정: {len(pending)}개")
    eta = len(pending) * ((DELAY_MIN + DELAY_MAX) / 2)
    print(f"예상 소요 시간: 약 {int(eta // 3600)}시간 {int((eta % 3600) // 60)}분\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        context = await make_context(browser)
        page = await context.new_page()

        for idx, (kw, bid) in enumerate(pending):
            overall = already_done + idx + 1
            print(f"[{overall:>5}/{actual_count}] ...", end=" ", flush=True)

            result = await search_blog_rank(page, kw, bid)
            mark_result(df, result)
            save_dataframe(df, output_path)

            if result["error"]:
                print(f"오류: {result['error']}")
            else:
                print(_brief(result["sections"]))

            if idx == len(pending) - 1:
                break

            if (idx + 1) % BATCH_SIZE == 0:
                pause = random.uniform(BATCH_BREAK_MIN, BATCH_BREAK_MAX)
                print(f"\n  {BATCH_SIZE}개 완료 — {pause:.0f}초 휴식...\n")
                await asyncio.sleep(pause)
                if (idx + 1) % CONTEXT_RESET_EVERY == 0:
                    await context.close()
                    context = await make_context(browser)
                    page = await context.new_page()
            else:
                await asyncio.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

        await browser.close()

    done_df = df[df["처리완료"] == "Y"]
    print(f"\n결과 저장: {output_path}")
    print(f"\n--- 요약 ({len(done_df)}/{actual_count}) ---")
    for sec in TRACKED_SECTIONS:
        cnt = done_df[sec].astype(str).str.match(r"^\d+").sum() if sec in done_df.columns else 0
        if cnt:
            print(f"  {sec}: {cnt}개 노출")


# ── Google Sheets 지원 ────────────────────────────────────────────────────────

from sheets_io import _api_call as _sh, _get_client as _sh_client  # noqa: E402


class BlogRankSheetsSession:
    """Google Sheets 블로그 순위 데이터 읽기/쓰기 세션."""

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

    def stage_result(self, row_idx: int, result: dict):
        """결과 한 건을 버퍼에 추가. FLUSH_EVERY에 도달하면 자동 flush."""
        for col_name, value in _build_updates(result).items():
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
        from collections import defaultdict
        from gspread.utils import a1_to_rowcol
        rows: dict = defaultdict(dict)
        for item in self._pending:
            row, col = a1_to_rowcol(item["range"])
            rows[row][col] = item["values"][0][0]
        for row_num in sorted(rows):
            col_vals = rows[row_num]
            min_col, max_col = min(col_vals), max(col_vals)
            values = [[col_vals.get(c, "") for c in range(min_col, max_col + 1)]]
            start = self._rowcol_to_a1(row_num, min_col)
            end = self._rowcol_to_a1(row_num, max_col)
            _sh(self._ws.update, f"{start}:{end}", values, value_input_option="RAW")
        self._pending = []
        self._staged_count = 0


async def run_sheets(spreadsheet_id: str, gid: int, headless: bool,
                     start: int, count: int | None, source_gid: int | None = None):
    """Google Sheets 모드: 소스 시트에서 읽고 결과 시트에 씀."""
    # source_gid 비면 결과 시트 자체에서 읽음 (create_blog_rank_sheet가 키워드/아이디를 복사해 둠)
    session = BlogRankSheetsSession(spreadsheet_id, gid, source_gid=source_gid or None)
    pairs, row_indices = session.read_pairs(start, count)

    if not pairs:
        print(f"범위(start={start})에 처리할 키워드/아이디 쌍이 없습니다.")
        return

    total = len(pairs)
    print(f"Google Sheets 모드 | 처리 예정: {total}개 (start={start})")
    eta = total * ((DELAY_MIN + DELAY_MAX) / 2)
    print(f"예상 소요 시간: 약 {int(eta // 3600)}시간 {int((eta % 3600) // 60)}분\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        context = await make_context(browser)
        page = await context.new_page()

        for idx, ((kw, bid), row_idx) in enumerate(zip(pairs, row_indices)):
            print(f"[{idx + 1:>5}/{total}] ...", end=" ", flush=True)

            result = await search_blog_rank(page, kw, bid)
            session.stage_result(row_idx, result)

            if result["error"]:
                print(f"오류: {result['error']}")
            else:
                print(_brief(result["sections"]))

            if idx == total - 1:
                break

            if (idx + 1) % BATCH_SIZE == 0:
                session.flush()
                pause = random.uniform(BATCH_BREAK_MIN, BATCH_BREAK_MAX)
                print(f"\n  {BATCH_SIZE}개 완료 — {pause:.0f}초 휴식...\n")
                await asyncio.sleep(pause)
                if (idx + 1) % CONTEXT_RESET_EVERY == 0:
                    await context.close()
                    context = await make_context(browser)
                    page = await context.new_page()
            else:
                await asyncio.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

        await browser.close()

    session.flush()
    print(f"\n완료: {total}개 처리 → 구글 시트에 저장됨")


def main():
    parser = argparse.ArgumentParser(description="키워드 검색 → 블로그 아이디 섹션별 순위 확인")
    parser.add_argument("input_file", nargs="?", help="키워드(A열)/아이디(B열) Excel 파일 경로")
    parser.add_argument("--start", type=int, default=0, help="시작 행 번호 (0-based, 기본: 0)")
    parser.add_argument("--count", type=int, default=None, help="처리할 키워드 수 (기본: 전체)")
    parser.add_argument("--output-dir", help="결과 파일 저장 폴더 (기본: 입력 파일과 동일)")
    parser.add_argument("--headless", action="store_true", help="브라우저 숨김 모드")
    # Google Sheets 모드
    parser.add_argument("--sheets-id", help="Google Spreadsheet ID")
    parser.add_argument("--sheet-gid", type=int, default=0, help="결과 시트 GID")
    parser.add_argument("--source-gid", type=int, default=None, help="키워드 소스 시트 GID (기본: --sheet-gid와 동일)")
    args = parser.parse_args()

    if args.sheets_id:
        asyncio.run(run_sheets(args.sheets_id, args.sheet_gid, args.headless,
                               args.start, args.count,
                               source_gid=args.source_gid or None))
    elif args.input_file:
        asyncio.run(run(args.input_file, args.headless, args.start, args.count, args.output_dir))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

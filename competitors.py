#!/usr/bin/env python3
"""
대륜 미노출 섹션 상위 1~3위 업체 확인

사용법:
    python competitors.py keywords.xlsx --headless
    python competitors.py keywords.xlsx --start 0 --count 100 --headless
"""

import asyncio
import argparse
import copy
import sys
import random
import os
from datetime import date
from pathlib import Path
from urllib.parse import quote

from playwright.async_api import async_playwright, Page

from naver_clip import MOBILE_UA, MOBILE_VIEWPORT, normalize_section, _get_target_keywords

DELAY_MIN = 2.0
DELAY_MAX = 5.0
BATCH_SIZE = 50
BATCH_BREAK_MIN = 15
BATCH_BREAK_MAX = 30
CONTEXT_RESET_EVERY = 200
NUM_WORKERS = 10


def _mask_for_ci(*values: str):
    """GitHub Actions 로그에서 값을 마스킹 (CI 환경에서만 동작)."""
    if not os.environ.get("GITHUB_ACTIONS"):
        return
    for v in values:
        if v:
            print(f"::add-mask::{v}", flush=True)

TRACKED_SECTIONS = [
    "네이버 클립", "뉴스", "인기글", "이미지", "웹문서 1", "웹문서 2", "플레이스",
]

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

# 대륜이 미노출인 섹션의 상위 1~3위 업체명까지 함께 반환하는 JS
_COMPETITOR_JS = """
(keywords) => {
    function hasTarget(html) {
        return keywords.some(k => html.includes(k));
    }
    function hasVisibleTarget(el) {
        const clone = el.cloneNode(true);
        clone.querySelectorAll('script, style').forEach(e => e.remove());
        return keywords.some(k => clone.innerHTML.includes(k));
    }
    function hasAltTarget(el) {
        return Array.from(el.querySelectorAll('img')).some(img =>
            keywords.some(k => (img.getAttribute('alt') || '').includes(k))
        );
    }
    function getItemName(item, isPlace, isImage) {
        if (isImage) {
            const img = item.querySelector('img');
            return img ? (img.getAttribute('alt') || '').trim().slice(0, 80) : '';
        }
        if (isPlace) {
            const sels = ['[class*="place_bluelink"]', '[class*="LWxIZ"]', '.YzBgS', 'a strong', 'strong'];
            for (const sel of sels) {
                const el = item.querySelector(sel);
                if (el && el.textContent.trim()) return el.textContent.trim().slice(0, 80);
            }
        }
        const titleSels = ['strong.title', '.title', 'a strong', '[class*="title"]', 'h3', 'h4', 'strong'];
        for (const sel of titleSels) {
            const el = item.querySelector(sel);
            if (el && el.textContent.trim()) return el.textContent.trim().slice(0, 80);
        }
        const lines = item.textContent.trim().split(/[\\n\\r\\t]/).map(l => l.trim()).filter(Boolean);
        return (lines[0] || '').slice(0, 80);
    }

    const results = [];

    const namedEls = Array.from(
        document.querySelectorAll('[class*="api_subject_bx"], .sc_new, [class*="sc_new"]')
    ).filter(el => {
        const h2 = el.querySelector('h2');
        return h2 && h2.textContent.trim();
    });

    function d2children(root) {
        const out = [];
        for (const c of root.children)
            for (const cc of c.children) out.push(cc);
        return out;
    }

    for (const sec of namedEls) {
        const h2 = sec.querySelector('h2');
        const name = h2.textContent.trim().replace(/\\s+/g, ' ');

        const d2 = d2children(sec);
        if (d2.length < 2) {
            results.push({ name, has_target: hasVisibleTarget(sec), position: null, total: 0, top_items: [] });
            continue;
        }

        const contentEl = d2[1];
        let items = Array.from(contentEl.children);
        if (items.length <= 2) {
            const gc = [];
            for (const c of contentEl.children)
                for (const cc of c.children) gc.push(cc);
            if (gc.length > items.length) items = gc;
        }
        items = items.filter(i => (i.textContent || '').trim().length > 0);

        const isImageSection = name.includes('이미지');
        const positions = [];
        for (let i = 0; i < items.length; i++) {
            const matched = isImageSection
                ? hasAltTarget(items[i])
                : hasTarget(items[i].innerHTML || '');
            if (matched) positions.push(i + 1);
        }

        const top_items = items.slice(0, 3).map(item => getItemName(item, false, isImageSection));

        results.push({
            name,
            has_target: positions.length > 0 || (items.length === 0 && (isImageSection ? hasAltTarget(sec) : hasVisibleTarget(sec))),
            position: positions[0] || null,
            positions,
            total: items.length,
            top_items,
        });
    }

    // 웹문서 섹션
    const webDocEntries = [];

    document.querySelectorAll('.spw_fsolid').forEach(sec => {
        const list = sec.querySelector('.fsolid_list');
        const items = list
            ? Array.from(list.children).filter(el => el.tagName === 'DIV')
            : Array.from(sec.children).filter(el => el.tagName === 'DIV');
        if (items.length > 0) webDocEntries.push({ el: sec, items });
    });

    document.querySelectorAll('[class*="fds-web-list-root"]').forEach(root => {
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
            if (hasTarget(items[i].textContent || '')) positions.push(i + 1);
        }
        const top_items = items.slice(0, 3).map(item => getItemName(item, false, false));
        results.push({
            name: `웹문서 ${idx + 1}`,
            has_target: positions.length > 0,
            position: positions[0] || null,
            positions,
            total: items.length,
            top_items,
        });
    });

    // 플레이스 섹션
    const placeSection = document.querySelector('[class*="place_section"]');
    if (placeSection) {
        let items = [];
        const ul = placeSection.querySelector('ul');
        if (ul) {
            items = Array.from(ul.children).filter(el => el.tagName === 'LI');
        }
        if (items.length === 0) {
            const container = placeSection.querySelector('[class*="place_list"], [class*="place_lst"]');
            if (container) items = Array.from(container.children);
        }
        if (items.length === 0) {
            items = Array.from(placeSection.querySelectorAll('[class*="place_item"], [class*="UEzoS"]'));
        }

        const positions = [];
        for (let i = 0; i < items.length; i++) {
            if (hasTarget(items[i].textContent || '')) positions.push(i + 1);
        }
        const has_target = positions.length > 0 || (items.length === 0 && hasTarget(placeSection.textContent || ''));
        const top_items = items.slice(0, 3).map(item => getItemName(item, true, false));
        results.push({
            name: '플레이스',
            has_target,
            position: positions[0] || null,
            positions,
            total: items.length,
            top_items,
        });
    }

    return results;
}
"""


def _build_result_cols() -> list[str]:
    cols = ["처리완료"]
    for sec in TRACKED_SECTIONS:
        cols += [f"{sec} 대륜", f"{sec} 1위", f"{sec} 2위", f"{sec} 3위"]
    cols.append("오류")
    return cols


RESULT_COLS = _build_result_cols()


async def search_competitors(page: Page, keyword: str) -> dict:
    result = {"keyword": keyword, "sections": [], "error": None}
    try:
        url = f"https://m.search.naver.com/search.naver?query={quote(keyword)}"
        await page.goto(url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(2_000)

        raw = await page.evaluate(_COMPETITOR_JS, _get_target_keywords())
        for sec in raw:
            sec["name"] = normalize_section(sec["name"])
        result["sections"] = raw

    except Exception as e:
        result["error"] = str(e)

    return result


def load_keywords(filepath: str, col: str | None, start: int, count: int | None) -> list[str]:
    import pandas as pd
    path = Path(filepath)
    if not path.exists():
        sys.exit(f"파일을 찾을 수 없습니다: {filepath}")

    xl = pd.ExcelFile(path)
    df = xl.parse(xl.sheet_names[0], dtype=str)
    if df.empty:
        sys.exit("엑셀 파일이 비어있습니다.")

    if col and col in df.columns:
        kw_col = col
    elif "키워드" in df.columns:
        kw_col = "키워드"
    else:
        kw_col = df.columns[0]

    all_kws = [k for k in df[kw_col].dropna().astype(str).str.strip().tolist() if k]
    return all_kws[start : (start + count) if count else None]


def load_or_create_output(output_path: Path, keywords: list[str]):
    import pandas as pd
    if output_path.exists():
        df = pd.read_excel(output_path, dtype=str)
        for c in RESULT_COLS:
            if c not in df.columns:
                df[c] = ""
        return df

    df = pd.DataFrame({"키워드": keywords})
    for c in RESULT_COLS:
        df[c] = ""
    return df


def save_dataframe(df, filepath: Path):
    import pandas as pd
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="경쟁사")
        ws = writer.sheets["경쟁사"]
        for col_cells in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col_cells), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 50)


def is_done(val) -> bool:
    import pandas as pd
    return pd.notna(val) and str(val).strip() == "Y"


def mark_result(df, result: dict):
    mask = df["키워드"].astype(str).str.strip() == result["keyword"]
    idxs = df[mask].index
    if not len(idxs):
        return
    i = idxs[0]

    for col in RESULT_COLS:
        df.at[i, col] = ""

    df.at[i, "오류"] = result["error"] or ""

    sec_map = {sec["name"]: sec for sec in result.get("sections", [])}

    for sec_name in TRACKED_SECTIONS:
        sec = sec_map.get(sec_name)
        if not sec:
            continue

        has = sec["has_target"]
        pos = sec.get("position")
        positions = sec.get("positions") or ([pos] if pos else [])
        top_items: list[str] = sec.get("top_items") or []

        if has:
            df.at[i, f"{sec_name} 대륜"] = (",".join(map(str, positions)) + "위") if positions else "있음"
            # 대륜이 노출 중일 때는 경쟁사 칸 비움
        else:
            df.at[i, f"{sec_name} 대륜"] = "X"
            for rank, name_text in enumerate(top_items[:3], start=1):
                df.at[i, f"{sec_name} {rank}위"] = name_text

    df.at[i, "처리완료"] = "Y"


def _brief(sections: list[dict]) -> str:
    found = []
    for s in sections:
        if s["has_target"]:
            positions = s.get("positions") or ([s["position"]] if s["position"] else [])
            if positions:
                found.append(f"{s['name']} {','.join(map(str, positions))}위")
            else:
                found.append(f"{s['name']} 있음")
    return " / ".join(found) if found else "미노출"


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


async def _worker_excel(worker_id: int, browser, queue: asyncio.Queue,
                        df, lock: asyncio.Lock, output_path,
                        counter: dict, total: int, already_done: int):
    await asyncio.sleep(worker_id * random.uniform(0.3, 0.7))
    context = await make_context(browser)
    page = await context.new_page()
    local_count = 0

    try:
        while True:
            try:
                kw = queue.get_nowait()
            except asyncio.QueueEmpty:
                break

            async with lock:
                counter["n"] += 1
                overall = already_done + counter["n"]

            print(f"[{overall:>5}/{total}] W{worker_id} ...", end=" ", flush=True)

            try:
                result = await search_competitors(page, kw)
                async with lock:
                    mark_result(df, result)
                    save_dataframe(df, output_path)
                if result["error"]:
                    print(f"오류: {result['error']}")
                else:
                    print(_brief(result["sections"]))
            except Exception as e:
                print(f"처리 오류 (skip): {e}")

            local_count += 1

            if local_count % BATCH_SIZE == 0:
                pause = random.uniform(BATCH_BREAK_MIN, BATCH_BREAK_MAX)
                print(f"\n  [W{worker_id}] {BATCH_SIZE}개 완료 — {pause:.0f}초 휴식...\n")
                await asyncio.sleep(pause)
                if local_count % CONTEXT_RESET_EVERY == 0:
                    await context.close()
                    context = await make_context(browser)
                    page = await context.new_page()
            else:
                await asyncio.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
    finally:
        await context.close()


async def run(input_file: str, col: str | None, headless: bool,
              start: int, count: int | None, output_dir: str | None,
              workers: int = NUM_WORKERS):

    keywords = load_keywords(input_file, col, start, count)
    if not keywords:
        sys.exit(f"범위(start={start})에 해당하는 키워드가 없습니다.")

    for kw in keywords:
        _mask_for_ci(kw)

    today = date.today().strftime("%Y%m%d")
    actual_count = len(keywords)
    out_dir = Path(output_dir) if output_dir else Path(input_file).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / f"competitors_{today}_{start}_{actual_count}.xlsx"

    df = load_or_create_output(output_path, keywords)

    done_mask = df["처리완료"].apply(is_done)
    already_done = int(done_mask.sum())
    pending_idx = df[~done_mask & df["키워드"].notna()].index.tolist()
    pending = [k for k in df.loc[pending_idx, "키워드"].astype(str).str.strip().tolist() if k]

    print(f"출력 파일: {output_path}")
    if already_done:
        print(f"이미 완료: {already_done}개 → 남은 {len(pending)}개 처리")
    if not pending:
        print("모든 키워드가 처리되었습니다.")
        return

    actual_workers = min(workers, len(pending))
    print(f"총 {actual_count}개 키워드 | 처리 예정: {len(pending)}개 | 워커: {actual_workers}개")
    eta = len(pending) * ((DELAY_MIN + DELAY_MAX) / 2) / actual_workers
    print(f"예상 소요 시간: 약 {int(eta // 3600)}시간 {int((eta % 3600) // 60)}분\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        queue: asyncio.Queue = asyncio.Queue()
        for kw in pending:
            queue.put_nowait(kw)

        lock = asyncio.Lock()
        counter = {"n": 0}

        tasks = [
            asyncio.create_task(
                _worker_excel(i, browser, queue, df, lock, output_path, counter, actual_count, already_done)
            )
            for i in range(actual_workers)
        ]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        await browser.close()
        for i, r in enumerate(results):
            if isinstance(r, Exception):
                print(f"\n[W{i}] 워커 오류: {r}")

    done_df = df[df["처리완료"] == "Y"]
    print(f"\n결과 저장: {output_path}")
    print(f"완료: {len(done_df)}/{actual_count}")


# ── Google Sheets 지원 ────────────────────────────────────────────────────────

from sheets_io import _api_call as _sh, _get_client as _sh_client, create_competitor_sheet  # noqa: E402


class CompetitorSheetsSession:
    """Google Sheets 경쟁사 데이터 읽기/쓰기 세션."""

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

    def read_keywords(self, start: int, count: int | None) -> tuple[list[str], list[int]]:
        src_all = _sh(self._source_ws.get_all_values)
        if len(src_all) < 2:
            return [], []
        src_header = src_all[0]
        kw_idx = next((i for i, h in enumerate(src_header) if h == "키워드"), 0)
        data_rows = src_all[1:]
        sliced = data_rows[start : (start + count) if count else None]

        done_col = self._header_map.get("처리완료")
        if self._separate_source:
            result_all = _sh(self._ws.get_all_values)
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
            row_indices.append(start + i + 2)

        return keywords, row_indices

    def stage_result(self, row_idx: int, result: dict):
        updates: dict[str, str] = {col: "" for col in RESULT_COLS}
        updates["오류"] = result["error"] or ""

        sec_map = {sec["name"]: sec for sec in result.get("sections", [])}
        for sec_name in TRACKED_SECTIONS:
            sec = sec_map.get(sec_name)
            if not sec:
                continue
            has = sec["has_target"]
            pos = sec.get("position")
            positions = sec.get("positions") or ([pos] if pos else [])
            top_items: list[str] = sec.get("top_items") or []
            if has:
                updates[f"{sec_name} 대륜"] = (",".join(map(str, positions)) + "위") if positions else "있음"
            else:
                updates[f"{sec_name} 대륜"] = "X"
                for rank, name_text in enumerate(top_items[:3], start=1):
                    updates[f"{sec_name} {rank}위"] = name_text

        updates["처리완료"] = "Y"

        for col_name, value in updates.items():
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
        from gspread.utils import a1_to_rowcol
        from collections import defaultdict
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


async def _worker_sheets(worker_id: int, browser, queue: asyncio.Queue,
                         session, lock: asyncio.Lock, counter: dict, total: int):
    await asyncio.sleep(worker_id * random.uniform(0.3, 0.7))
    context = await make_context(browser)
    page = await context.new_page()
    local_count = 0

    try:
        while True:
            try:
                kw, row_idx = queue.get_nowait()
            except asyncio.QueueEmpty:
                break

            async with lock:
                counter["n"] += 1
                overall = counter["n"]

            print(f"[{overall:>5}/{total}] W{worker_id} ...", end=" ", flush=True)

            try:
                result = await search_competitors(page, kw)
                async with lock:
                    session.stage_result(row_idx, result)
                if result["error"]:
                    print(f"오류: {result['error']}")
                else:
                    print(_brief(result["sections"]))
            except Exception as e:
                print(f"처리 오류 (skip): {e}")

            local_count += 1

            if local_count % BATCH_SIZE == 0:
                try:
                    async with lock:
                        session.flush()
                except Exception as e:
                    print(f"\n  [W{worker_id}] flush 오류: {e}")
                pause = random.uniform(BATCH_BREAK_MIN, BATCH_BREAK_MAX)
                print(f"\n  [W{worker_id}] {BATCH_SIZE}개 완료 — {pause:.0f}초 휴식...\n")
                await asyncio.sleep(pause)
                if local_count % CONTEXT_RESET_EVERY == 0:
                    await context.close()
                    context = await make_context(browser)
                    page = await context.new_page()
            else:
                await asyncio.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
    finally:
        await context.close()


async def run_sheets(spreadsheet_id: str, gid: int, headless: bool,
                     start: int, count: int | None, source_gid: int | None = None,
                     workers: int = NUM_WORKERS):
    session = CompetitorSheetsSession(spreadsheet_id, gid, source_gid=source_gid)
    keywords, row_indices = session.read_keywords(start, count)

    for kw in keywords:
        _mask_for_ci(kw)

    if not keywords:
        print(f"범위(start={start})에 처리할 키워드가 없습니다.")
        return

    total = len(keywords)
    actual_workers = min(workers, total)
    print(f"Google Sheets 모드 | 처리 예정: {total}개 (start={start}) | 워커: {actual_workers}개")
    eta = total * ((DELAY_MIN + DELAY_MAX) / 2) / actual_workers
    print(f"예상 소요 시간: 약 {int(eta // 3600)}시간 {int((eta % 3600) // 60)}분\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        queue: asyncio.Queue = asyncio.Queue()
        for kw, row_idx in zip(keywords, row_indices):
            queue.put_nowait((kw, row_idx))

        lock = asyncio.Lock()
        counter = {"n": 0}

        tasks = [
            asyncio.create_task(
                _worker_sheets(i, browser, queue, session, lock, counter, total)
            )
            for i in range(actual_workers)
        ]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        await browser.close()
        for i, r in enumerate(results):
            if isinstance(r, Exception):
                print(f"\n[W{i}] 워커 오류: {r}")

    session.flush()
    print(f"\n완료: {total}개 처리 → 구글 시트에 저장됨")


def main():
    parser = argparse.ArgumentParser(description="대륜 미노출 섹션 경쟁사 상위 1~3위 확인")
    parser.add_argument("input_file", nargs="?", help="키워드 Excel 파일 경로")
    parser.add_argument("--col", help="키워드 컬럼명")
    parser.add_argument("--start", type=int, default=0, help="시작 행 번호 (0-based)")
    parser.add_argument("--count", type=int, default=None, help="처리할 키워드 수")
    parser.add_argument("--output-dir", help="결과 파일 저장 폴더")
    parser.add_argument("--headless", action="store_true", help="브라우저 숨김 모드")
    # Google Sheets 모드
    parser.add_argument("--sheets-id", help="Google Spreadsheet ID")
    parser.add_argument("--sheet-gid", type=int, default=0, help="결과 시트 GID")
    parser.add_argument("--source-gid", type=int, default=None, help="키워드 소스 시트 GID")
    parser.add_argument("--workers", type=int, default=NUM_WORKERS, help=f"동시 실행 워커 수 (기본: {NUM_WORKERS})")
    args = parser.parse_args()

    if args.sheets_id:
        asyncio.run(run_sheets(args.sheets_id, args.sheet_gid, args.headless,
                               args.start, args.count, source_gid=args.source_gid,
                               workers=args.workers))
    elif args.input_file:
        asyncio.run(run(args.input_file, args.col, args.headless,
                        args.start, args.count, args.output_dir, workers=args.workers))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

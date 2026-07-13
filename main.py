#!/usr/bin/env python3
"""
네이버 모바일 검색결과 전체 섹션 순위 확인

사용법:
    # 단일 PC (전체 키워드)
    python main.py keywords.xlsx --headless

    # 다중 PC 분산 실행 (PC당 범위 지정)
    python main.py keywords.xlsx --start 0    --count 500 --output-dir ./results --headless
    python main.py keywords.xlsx --start 500  --count 500 --output-dir ./results --headless

    # 구조 확인
    python main.py --inspect "변호사 상담"
"""

import asyncio
import argparse
import sys
import random
from datetime import date
from pathlib import Path

import pandas as pd
from playwright.async_api import async_playwright

from naver_clip import search_all_sections, run_inspection, MOBILE_UA, MOBILE_VIEWPORT
from sheets_io import SheetsSession

DELAY_MIN = 2.0
DELAY_MAX = 5.0
BATCH_SIZE = 50
BATCH_BREAK_MIN = 15
BATCH_BREAK_MAX = 30
CONTEXT_RESET_EVERY = 200

TRACKED_SECTIONS = [
    "네이버 클립", "뉴스", "웹문서 1", "웹문서 2", "플레이스",
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


def load_keywords(filepath: str, col: str | None, start: int, count: int | None) -> tuple[list[str], str]:
    """소스 파일에서 키워드 목록만 읽기 (소스 파일은 수정하지 않음)."""
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
    sliced = all_kws[start : (start + count) if count else None]
    return sliced, kw_col


def load_or_create_output(output_path: Path, keywords: list[str]) -> pd.DataFrame:
    """결과 파일 로드(재시작 시) 또는 신규 생성."""
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


def save_dataframe(df: pd.DataFrame, filepath: Path):
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="결과")
        ws = writer.sheets["결과"]
        for col_cells in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col_cells), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 40)


def mark_result(df: pd.DataFrame, result: dict):
    mask = df["키워드"].astype(str).str.strip() == result["keyword"]
    idxs = df[mask].index
    if len(idxs) == 0:
        return
    i = idxs[0]

    for col in TRACKED_SECTIONS:
        df.at[i, col] = ""
    df.at[i, "기타 노출"] = ""
    df.at[i, "오류"] = result["error"] or ""

    other_found: list[str] = []
    for sec in result.get("sections", []):
        name: str = sec["name"]
        has: bool = sec["has_target"]
        pos: int | None = sec["position"]

        cell_val = (str(pos) if pos else "있음") if has else "X"

        if name in TRACKED_SECTIONS:
            df.at[i, name] = cell_val
        elif has:
            other_found.append(f"{name}:{cell_val}")

    df.at[i, "기타 노출"] = ", ".join(other_found)
    df.at[i, "처리완료"] = "Y"


def is_done(val) -> bool:
    return pd.notna(val) and str(val).strip() == "Y"


def _brief(sections: list[dict]) -> str:
    found = [
        f"{s['name']} {s['position']}위" if s["position"] else f"{s['name']} 있음"
        for s in sections if s["has_target"]
    ]
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


async def run(input_file: str, col: str | None, headless: bool,
              start: int, count: int | None, output_dir: str | None):

    keywords, _ = load_keywords(input_file, col, start, count)
    if not keywords:
        sys.exit(f"범위(start={start})에 해당하는 키워드가 없습니다.")

    today = date.today().strftime("%Y%m%d")
    actual_count = len(keywords)
    out_dir = Path(output_dir) if output_dir else Path(input_file).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / f"result_{today}_{start}_{actual_count}.xlsx"

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

    print(f"총 {actual_count}개 키워드 | 처리 예정: {len(pending)}개")
    eta = len(pending) * ((DELAY_MIN + DELAY_MAX) / 2)
    print(f"예상 소요 시간: 약 {int(eta // 3600)}시간 {int((eta % 3600) // 60)}분\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        context = await make_context(browser)
        page = await context.new_page()

        for idx, kw in enumerate(pending):
            overall = already_done + idx + 1
            print(f"[{overall:>5}/{actual_count}] {kw!r} ...", end=" ", flush=True)

            result = await search_all_sections(page, kw)
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
        cnt = done_df[sec].astype(str).str.match(r"^\d+$").sum() if sec in done_df.columns else 0
        if cnt:
            print(f"  {sec}: {cnt}개 노출")


async def run_sheets(spreadsheet_id: str, gid: int, headless: bool,
                     start: int, count: int | None, source_gid: int | None = None):
    """Google Sheets 모드: 소스 시트에서 읽고 결과 시트에 씀."""
    session = SheetsSession(spreadsheet_id, gid, source_gid=source_gid)
    keywords, row_indices = session.read_keywords(start, count)

    if not keywords:
        print(f"범위(start={start})에 처리할 키워드가 없습니다.")
        return

    total = len(keywords)
    print(f"Google Sheets 모드 | 처리 예정: {total}개 (start={start})")
    eta = total * ((DELAY_MIN + DELAY_MAX) / 2)
    print(f"예상 소요 시간: 약 {int(eta // 3600)}시간 {int((eta % 3600) // 60)}분\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        context = await make_context(browser)
        page = await context.new_page()

        for idx, (kw, row_idx) in enumerate(zip(keywords, row_indices)):
            print(f"[{idx + 1:>5}/{total}]", end=" ", flush=True)

            result = await search_all_sections(page, kw)
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
    parser = argparse.ArgumentParser(description="네이버 모바일 검색 순위 확인")
    parser.add_argument("input_file", nargs="?", help="키워드 Excel 파일 경로")
    parser.add_argument("--col", help="키워드 컬럼명")
    parser.add_argument("--start", type=int, default=0, help="시작 행 번호 (0-based, 기본: 0)")
    parser.add_argument("--count", type=int, default=None, help="처리할 키워드 수 (기본: 전체)")
    parser.add_argument("--output-dir", help="결과 파일 저장 폴더 (기본: 입력 파일과 동일)")
    parser.add_argument("--headless", action="store_true", help="브라우저 숨김 모드")
    parser.add_argument("--inspect", metavar="KEYWORD", help="단일 키워드 섹션 구조 확인")
    # Google Sheets 모드
    parser.add_argument("--sheets-id", help="Google Spreadsheet ID")
    parser.add_argument("--sheet-gid", type=int, default=0, help="결과 시트 GID")
    parser.add_argument("--source-gid", type=int, default=None, help="키워드 소스 시트 GID (기본: --sheet-gid와 동일)")
    parser.add_argument("--count-only", action="store_true",
                        help="키워드 수만 출력하고 종료 (GitHub Actions prepare 용)")
    args = parser.parse_args()

    if args.inspect:
        asyncio.run(run_inspection(args.inspect))
    elif args.sheets_id:
        if args.count_only:
            session = SheetsSession(args.sheets_id, args.sheet_gid, source_gid=args.source_gid)
            print(session.count_keywords())
        else:
            asyncio.run(run_sheets(args.sheets_id, args.sheet_gid, args.headless,
                                   args.start, args.count, source_gid=args.source_gid))
    elif args.input_file:
        asyncio.run(run(args.input_file, args.col, args.headless,
                        args.start, args.count, args.output_dir))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

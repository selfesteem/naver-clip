#!/usr/bin/env python3
"""
네이버 클립 섹션 순위 검출 — 테스트 래퍼 (clip_rank 엔진 사용)

소규모 검증용: 키워드 직접 지정 / 시트 앞 N개 실행 + 로컬 xlsx 백업.
시트 쓰기는 clip_rank 와 동일하게 '소스 행 번호 = 결과 행 번호' 규칙을 따름.

사용법:
    python clip_rank_test.py                                    # 기본 테스트 키워드 2개
    python clip_rank_test.py --keyword 부동산법무법인
    python clip_rank_test.py --keyword "부동산법무법인,공무집행방해변호사"
    python clip_rank_test.py --limit 5                          # 시트 앞 5개 키워드
    python clip_rank_test.py --no-sheets                        # 시트 쓰기 생략(로컬 xlsx만)
    python clip_rank_test.py --headed                           # 브라우저 보이게

주의: 워커 실행 중에는 로컬 시트 쓰기 금지 (동일 행 충돌) — 그때는 --no-sheets 사용.
"""

import asyncio
import argparse
import random
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from playwright.async_api import async_playwright

load_dotenv()

from clip_rank import (
    DELAY_MAX,
    DELAY_MIN,
    RESULT_COLS,
    ResultWriter,
    _brief,
    _env_credentials,
    _result_row,
    _spreadsheet_id,
    _summary,
    ensure_result_tab,
    load_source,
    make_context,
    search_clip_rank,
)

DEFAULT_TEST_KEYWORDS = ["공무집행방해변호사", "부동산법무법인"]


def save_local(results: list[dict]) -> Path:
    out_dir = Path("results")
    out_dir.mkdir(exist_ok=True)
    out = out_dir / f"clip_rank_test_{datetime.now():%Y%m%d_%H%M}.xlsx"
    df = pd.DataFrame([_result_row(r) for r in results], columns=RESULT_COLS)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="클립순위")
        ws = writer.sheets["클립순위"]
        for col_cells in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col_cells), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 45)
    return out


async def run(targets: list[tuple[int | None, str]], channel_ids: set[str],
              headless: bool, use_sheets: bool):
    """targets: [(소스 행 번호, 키워드)] — 행 번호 None이면 로컬 기록만."""
    total = len(targets)
    print(f"처리 예정: {total}개 키워드\n")

    writer = None
    tab = ""
    if use_sheets:
        sid = _spreadsheet_id(None)
        row_keywords = {row: kw for row, kw in targets if row is not None}
        ws, tab = ensure_result_tab(sid, row_keywords)
        writer = ResultWriter(ws)

    results: list[dict] = []
    staged = 0

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        context = await make_context(browser)
        page = await context.new_page()

        for idx, (row, kw) in enumerate(targets):
            label = f"{row}행" if row is not None else "시트밖"
            print(f"[{idx + 1:>3}/{total}] {label} {kw} ...", end=" ", flush=True)

            r = await search_clip_rank(page, kw, channel_ids)
            results.append(r)
            if row is not None and writer is not None:
                writer.stage(row, r)
                staged += 1
            print(_brief(r))

            if idx < total - 1:
                await asyncio.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

        await browser.close()

    if writer is not None:
        writer.flush()
        print(f"\n시트 저장: {staged}행 → '{tab}'")

    local_path = save_local(results)
    print(f"로컬 저장: {local_path}")
    print(f"\n--- 요약 ---\n  {_summary(results)}")


def main():
    parser = argparse.ArgumentParser(description="네이버 클립 섹션 순위 검출 (테스트)")
    parser.add_argument("--keyword", help="직접 지정 키워드 (쉼표 구분). 생략 시 기본 테스트 키워드")
    parser.add_argument("--limit", type=int, default=None, help="시트 키워드 중 앞 N개만 처리")
    parser.add_argument("--no-sheets", action="store_true", help="구글 시트 저장 생략")
    parser.add_argument("--headed", action="store_true", help="브라우저 창 표시")
    args = parser.parse_args()

    _env_credentials()
    row_keywords, channel_ids = load_source(_spreadsheet_id(None))
    print(f"시트 로드: 키워드 {len(row_keywords)}개 | 채널 {len(channel_ids)}개 (1:N 매칭용)")

    kw_to_row = {kw: row for row, kw in row_keywords.items()}

    if args.keyword:
        kws = [k.strip() for k in args.keyword.split(",") if k.strip()]
    elif args.limit is not None:
        kws = [kw for _, kw in sorted(row_keywords.items())][: args.limit]
    else:
        kws = DEFAULT_TEST_KEYWORDS

    targets: list[tuple[int | None, str]] = []
    for kw in kws:
        row = kw_to_row.get(kw)
        if row is None:
            print(f"참고: 시트에 없는 키워드 → 로컬 기록만: {kw}")
        targets.append((row, kw))

    asyncio.run(run(targets, channel_ids, headless=not args.headed, use_sheets=not args.no_sheets))


if __name__ == "__main__":
    main()

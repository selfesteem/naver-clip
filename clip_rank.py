#!/usr/bin/env python3
"""
네이버 모바일 검색 '네이버 클립' 섹션 채널 노출 순위 검출 (스케쥴러/워커용)

동작 (키워드당):
  1) 모바일 네이버에서 키워드 검색
  2) '네이버 클립' 섹션 확인 — 없으면 패스(섹션없음)
  3) 클립 섹션의 클립을 DOM 순서대로 수집 후 채널 3조건 매칭:
     (1) 채널명에 "대륜" 포함
     (2) 채널 URL에 "daeryun" 포함
     (3) 시트 '채널' 열에 등록된 URL 중 하나와 채널 ID 일치 (1:N)
  4) 매칭된 클립의 순위를 날짜별 탭(클립순위_MMDD)의 해당 행에 기록

행 규칙 (워커 충돌 방지 핵심):
  - 결과 탭 행 번호 = 소스 탭('키워드/채널') 행 번호 (1:1 고정, prepare가 프리필)
  - 워커는 배정받은 행(--row-indices)만 갱신 — 다른 행 절대 쓰지 않음
  - 처리완료=Y인 행은 스킵 → 재실행 시 이어서 진행 (재시작 safe)

사용법:
    # 전체 실행 (미완료만)
    python clip_rank.py --headless

    # GitHub Actions 워커: 배정받은 행만
    python clip_rank.py --row-indices "2,5,10-20" --tab 클립순위_0921 --headless

    # prepare: 미완료 행을 N개 워커에 분할해 config.json 생성
    python clip_rank.py --prepare --workers 10

    # summary: 전체 워커 종료 후 탭 마지막 행에 노출 키워드 합계 기록
    python clip_rank.py --summary --tab 클립순위_0921

환경변수:
    GOOGLE_CREDENTIALS      서비스계정 JSON 문자열 (CI)
    GOOGLE_CREDENTIALS_JSON 서비스계정 JSON 파일 경로 (로컬, .env)
    GOOGLE_SHEET_ID / SPREADSHEET_ID  스프레드시트 ID (--sheets-id가 우선)
"""

import asyncio
import argparse
import copy
import json
import math
import os
import random
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

from dotenv import load_dotenv
from gspread import Worksheet
from playwright.async_api import async_playwright, Page

from naver_clip import MOBILE_UA
from sheets_io import _api_call, _get_client

load_dotenv()  # CI에는 .env 없음 → 무동작, 로컬 편의용 (기존 환경변수는 덮어쓰지 않음)

# ── 설정 ─────────────────────────────────────────────────────────

KST = timezone(timedelta(hours=9))

NAME_KEYWORDS = ["대륜"]          # 조건1: 채널명 매칭 키워드
URL_KEYWORDS = ["daeryun"]        # 조건2: 채널 URL 매칭 키워드

SOURCE_TAB = "키워드/채널"         # A열 키워드, B열 채널 URL
RESULT_TAB_PREFIX = "클립순위_"    # 날짜별 탭: 클립순위_0921 (KST 기준)

RESULT_COLS = ["키워드", "순위", "매칭채널", "매칭조건", "클립제목", "총클립수", "처리완료", "오류"]
COL_DONE = 7                      # G열 = 처리완료

FLUSH_EVERY = 10                  # 결과 N개마다 시트 batch_update
DELAY_MIN = 2.0
DELAY_MAX = 5.0
BATCH_SIZE = 50
BATCH_BREAK_MIN = 15
BATCH_BREAK_MAX = 30
CONTEXT_RESET_EVERY = 200

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

# tv.naver.com/{id} 에서 채널 ID가 아닌 예약 경로 (영상/검색 링크 등)
RESERVED_TV_PATHS = {"v", "search", "clips", "my", "feed", "popular", "ranking"}

# ── 클립 섹션 수집 JS ─────────────────────────────────────────────
# 반환: { section_found, clips: [{rank, channel_id, href, channel_text, title}] }

_CLIP_JS = """
() => {
    function findClipSection() {
        for (const el of document.querySelectorAll(
            '[class*="api_subject_bx"], .sc_new, [class*="sc_new"]'
        )) {
            const h2 = el.querySelector('h2');
            if (h2 && h2.textContent.includes('클립')) return el;
        }
        return null;
    }

    const clipSec = findClipSection();
    if (!clipSec) return { section_found: false, clips: [] };

    const RESERVED = new Set(['v', 'search', 'clips', 'my', 'feed', 'popular', 'ranking']);
    const anchors = Array.from(clipSec.querySelectorAll('a[href]'));
    const clips = [];
    let lastTitle = '';

    for (const a of anchors) {
        const href = a.getAttribute('href') || '';
        const text = (a.textContent || '').trim().replace(/\\s+/g, ' ');

        let channelId = null;
        let m = href.match(/clip\\.naver\\.com\\/@([A-Za-z0-9_-]+)/);
        if (m) {
            channelId = m[1];
        } else {
            m = href.match(/tv\\.naver\\.com\\/([A-Za-z0-9_-]+)/);
            if (m && !RESERVED.has(m[1])) channelId = m[1];
        }

        if (channelId) {
            const prev = clips[clips.length - 1];
            if (prev && prev.channel_id === channelId && !prev.title) {
                prev.title = lastTitle;  // 같은 클립의 보조 링크 → 병합
            } else {
                clips.push({
                    rank: clips.length + 1,
                    channel_id: channelId,
                    href,
                    channel_text: text.slice(0, 80),
                    title: lastTitle,
                });
            }
        } else if (href.includes('m.naver.com/shorts') && text) {
            lastTitle = text.slice(0, 80);
        }
    }
    return { section_found: true, clips };
}
"""

# 클립 캐러셀 가로 스크롤 → 지연 로딩된 추가 클립 확보 (best-effort)
_SCROLL_JS = """
() => {
    function findClipSection() {
        for (const el of document.querySelectorAll(
            '[class*="api_subject_bx"], .sc_new, [class*="sc_new"]'
        )) {
            const h2 = el.querySelector('h2');
            if (h2 && h2.textContent.includes('클립')) return el;
        }
        return null;
    }
    const sec = findClipSection();
    if (!sec) return 0;
    let scrolled = 0;
    for (const el of sec.querySelectorAll('*')) {
        const st = getComputedStyle(el);
        if ((st.overflowX === 'auto' || st.overflowX === 'scroll') && el.scrollWidth > el.clientWidth) {
            el.scrollLeft = el.scrollWidth;
            scrolled += 1;
        }
    }
    return scrolled;
}
"""


# ── 채널 ID 추출 / 매칭 ──────────────────────────────────────────

def extract_channel_id(url: str) -> str | None:
    """채널 URL에서 채널 ID 추출.

    https://tv.naver.com/drobstructdutydr      → drobstructdutydr
    https://clip.naver.com/@daeryun-lawyer     → daeryun-lawyer
    """
    url = url.strip()
    if not url:
        return None
    m = re.search(r"clip\.naver\.com/@([A-Za-z0-9_-]+)", url)
    if m:
        return m.group(1)
    m = re.search(r"tv\.naver\.com/([A-Za-z0-9_-]+)", url)
    if m and m.group(1) not in RESERVED_TV_PATHS:
        return m.group(1)
    return None


def match_clip(clip: dict, sheet_channel_ids: set[str]) -> list[str]:
    """클립 1개에 대해 매칭된 조건 목록 반환 (없으면 빈 리스트)."""
    conds: list[str] = []
    if any(k in clip["channel_text"] for k in NAME_KEYWORDS):
        conds.append("채널명")
    if any(k in clip["href"].lower() for k in URL_KEYWORDS):
        conds.append("URL")
    if clip["channel_id"] in sheet_channel_ids:
        conds.append("시트채널")
    return conds


# ── 검색 엔진 ─────────────────────────────────────────────────────

async def search_clip_rank(page: Page, keyword: str, sheet_channel_ids: set[str]) -> dict:
    """키워드 1개 검색 → 클립 섹션 매칭 결과 반환."""
    result: dict = {
        "keyword": keyword, "status": "ok", "ranks": [], "matched": [],
        "total_clips": 0, "error": None,
    }
    try:
        url = f"https://m.search.naver.com/search.naver?query={quote(keyword)}"
        await page.goto(url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(2_000)

        raw = await page.evaluate(_CLIP_JS)

        # 캐러셀 스크롤로 추가 클립 로딩 시도 → 더 많이 수집된 쪽 사용
        if raw["section_found"]:
            scrolled = await page.evaluate(_SCROLL_JS)
            if scrolled > 0:
                await page.wait_for_timeout(1_500)
                raw2 = await page.evaluate(_CLIP_JS)
                if len(raw2["clips"]) > len(raw["clips"]):
                    raw = raw2

        if not raw["section_found"]:
            result["status"] = "섹션없음"
            return result

        clips = raw["clips"]
        result["total_clips"] = len(clips)

        for clip in clips:
            conds = match_clip(clip, sheet_channel_ids)
            if conds:
                result["ranks"].append(clip["rank"])
                result["matched"].append({
                    "rank": clip["rank"],
                    "channel_id": clip["channel_id"],
                    "conds": "+".join(conds),
                    "title": clip["title"],
                })

        if clips and not result["matched"]:
            result["status"] = "미노출"

    except Exception as e:
        result["error"] = str(e)
        result["status"] = "오류"
    return result


def _result_row(r: dict) -> list[str]:
    """결과 dict → 시트 행. 순위: '3,5' | X | 섹션없음."""
    if r["status"] == "섹션없음":
        rank_val, chans, conds, titles = "섹션없음", "", "", ""
    elif r["status"] == "오류":
        rank_val, chans, conds, titles = "", "", "", ""
    elif r["matched"]:
        rank_val = ",".join(str(m["rank"]) for m in r["matched"])
        chans = ", ".join(m["channel_id"] for m in r["matched"])
        conds = " | ".join(m["conds"] for m in r["matched"])
        titles = " / ".join(m["title"] for m in r["matched"])
    else:
        rank_val, chans, conds, titles = "X", "", "", ""
    return [
        r["keyword"], rank_val, chans, conds, titles,
        str(r["total_clips"]), "Y" if r["status"] != "오류" else "", r["error"] or "",
    ]


def _mask_channel(channel_id: str) -> str:
    """CI(공개 저장소) 로그에서 채널 ID 비식별화. 로컬은 그대로 출력."""
    return "***" if os.environ.get("GITHUB_ACTIONS") else channel_id


def _brief(r: dict) -> str:
    if r["status"] == "섹션없음":
        return "클립 섹션 없음 → 패스"
    if r["status"] == "오류":
        return f"오류: {r['error']}"
    if r["matched"]:
        parts = [f"{m['rank']}위({_mask_channel(m['channel_id'])}, {m['conds']})" for m in r["matched"]]
        return f"★ {' | '.join(parts)} / 총 {r['total_clips']}개"
    return f"미노출 (총 {r['total_clips']}개)"


# ── Google Sheets ─────────────────────────────────────────────────

def _env_credentials() -> None:
    """GOOGLE_CREDENTIALS(JSON 문자열) 없으면 GOOGLE_CREDENTIALS_JSON(파일 경로)에서 로드."""
    if os.environ.get("GOOGLE_CREDENTIALS"):
        return
    path = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
    if path and Path(path).exists():
        os.environ["GOOGLE_CREDENTIALS"] = Path(path).read_text(encoding="utf-8")


def _spreadsheet_id(cli_id: str | None) -> str:
    sid = cli_id or os.environ.get("GOOGLE_SHEET_ID") or os.environ.get("SPREADSHEET_ID")
    if not sid:
        raise EnvironmentError("--sheets-id 또는 GOOGLE_SHEET_ID/SPREADSHEET_ID 가 필요합니다.")
    return sid


def load_source(spreadsheet_id: str) -> tuple[dict[int, str], set[str]]:
    """'키워드/채널' 탭 → ({행번호: 키워드}, 채널 ID 전체 집합(1:N))."""
    _env_credentials()
    client = _get_client()
    ss = client.open_by_key(spreadsheet_id)
    ws = next((w for w in _api_call(ss.worksheets) if w.title == SOURCE_TAB), None)
    if ws is None:
        raise ValueError(f"탭 '{SOURCE_TAB}' 을 찾을 수 없습니다.")

    vals = _api_call(ws.get_all_values)
    row_keywords: dict[int, str] = {}
    channel_ids: set[str] = set()
    for i, row in enumerate(vals[1:]):
        row_num = i + 2
        kw = row[0].strip() if row else ""
        if kw:
            row_keywords[row_num] = kw
        if len(row) > 1:
            cid = extract_channel_id(row[1])
            if cid:
                channel_ids.add(cid)
    return row_keywords, channel_ids


def result_tab_name(now: datetime | None = None) -> str:
    return f"{RESULT_TAB_PREFIX}{(now or datetime.now(KST)):%m%d}"


def ensure_result_tab(
    spreadsheet_id: str, row_keywords: dict[int, str], tab_name: str | None = None
) -> tuple[Worksheet, str]:
    """날짜별 결과 탭 확보. 없으면 생성 + 소스 행 번호와 1:1로 키워드 프리필."""
    client = _get_client()
    ss = client.open_by_key(spreadsheet_id)
    tab = tab_name or result_tab_name()
    ws = next((w for w in _api_call(ss.worksheets) if w.title == tab), None)
    if ws is not None:
        return ws, tab

    max_row = max(row_keywords)
    ws = _api_call(ss.add_worksheet, title=tab, rows=max_row + 1, cols=len(RESULT_COLS))
    _api_call(ws.update, [RESULT_COLS], "A1")
    grid = [[row_keywords.get(r, "")] for r in range(2, max_row + 1)]
    _api_call(ws.update, grid, "A2")
    print(f"새 탭 생성: '{tab}' (키워드 {len(row_keywords)}개 프리필, 소스 행과 1:1 정렬)")
    return ws, tab


def pending_rows(
    ws: Worksheet, row_keywords: dict[int, str], row_indices: list[int] | None = None
) -> list[int]:
    """처리완료≠Y 인 대상 행 (오름차순). row_indices 지정 시 그중 미완료만."""
    targets = sorted(row_keywords) if row_indices is None else sorted(set(row_indices))
    done_vals = _api_call(ws.col_values, COL_DONE)  # index 0 = 1행(헤더)
    done = {i + 1 for i, v in enumerate(done_vals) if v.strip() == "Y"}
    return [r for r in targets if r not in done]


def parse_row_indices(spec: str) -> list[int]:
    """'2,5,10-20' → [2,5,10,...,20] (정렬·중복 제거)."""
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part and not part.startswith("-"):
            a, b = part.split("-", 1)
            out.extend(range(int(a), int(b) + 1))
        elif part.isdigit():
            out.append(int(part))
        else:
            raise ValueError(f"잘못된 행 번호 표기: {part!r}")
    return sorted(set(out))


class ResultWriter:
    """배정받은 행만 갱신. FLUSH_EVERY마다 batch_update 1회(요청 단위 원자적)."""

    FLUSH_EVERY = 10

    def __init__(self, ws: Worksheet):
        self._ws = ws
        self._pending: list[dict] = []
        self.staged = 0

    def stage(self, row_idx: int, result: dict) -> None:
        self._pending.append({"range": f"A{row_idx}", "values": [_result_row(result)]})
        self.staged += 1
        if len(self._pending) >= self.FLUSH_EVERY:
            self.flush()

    def flush(self) -> None:
        if not self._pending:
            return
        _api_call(self._ws.batch_update, copy.deepcopy(self._pending), value_input_option="RAW")
        self._pending = []


# ── 실행 ──────────────────────────────────────────────────────────

def _mask_for_ci(*values: str):
    if os.environ.get("GITHUB_ACTIONS"):
        for v in values:
            if v:
                print(f"::add-mask::{v}", flush=True)


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


def _summary(results: list[dict]) -> str:
    return (
        f"노출: {sum(1 for r in results if r['matched'])}개 | "
        f"미노출: {sum(1 for r in results if r['status'] == '미노출')}개 | "
        f"섹션없음: {sum(1 for r in results if r['status'] == '섹션없음')}개 | "
        f"오류: {sum(1 for r in results if r['status'] == '오류')}개"
    )


async def run_worker(args) -> None:
    _mask_for_ci(*NAME_KEYWORDS, *URL_KEYWORDS)
    sid = _spreadsheet_id(args.sheets_id)
    row_keywords, channel_ids = load_source(sid)
    _mask_for_ci(*channel_ids)  # 채널 ID가 다른 로그 경로(에러 등)로 새는 것 방지
    print(f"시트 로드: 키워드 {len(row_keywords)}개 | 채널 {len(channel_ids)}개 (1:N 매칭용)")

    ws, tab = ensure_result_tab(sid, row_keywords, tab_name=args.tab)

    if args.row_indices:
        indices = parse_row_indices(args.row_indices)
    else:
        all_rows = sorted(row_keywords)
        end = args.start + args.count if args.count else None
        indices = all_rows[args.start:end]

    rows = pending_rows(ws, row_keywords, indices)
    total = len(rows)
    print(f"'{tab}' 처리 예정: {total}행 (요청 {len(indices)}행 중 미완료)")
    if not total:
        print("처리할 행이 없습니다.")
        return

    writer = ResultWriter(ws)
    results: list[dict] = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=args.headless)
        context = await make_context(browser)
        page = await context.new_page()

        for idx, row in enumerate(rows):
            kw = row_keywords[row]
            _mask_for_ci(kw)
            print(f"[{idx + 1:>4}/{total}] {row}행 {kw} ...", end=" ", flush=True)

            r = await search_clip_rank(page, kw, channel_ids)
            writer.stage(row, r)
            results.append(r)
            print(_brief(r))

            if idx == total - 1:
                break

            if (idx + 1) % BATCH_SIZE == 0:
                writer.flush()
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

    writer.flush()
    print(f"\n완료: {total}행 → '{tab}'")
    print(f"--- 요약 ---\n  {_summary(results)}")


# ── prepare (워커 배분) ───────────────────────────────────────────

def run_prepare(args) -> None:
    """날짜 탭 생성 + 미완료 행을 N 워커에 disjoint 분할해 config.json 생성."""
    sid = _spreadsheet_id(args.sheets_id)
    row_keywords, _ = load_source(sid)
    ws, tab = ensure_result_tab(sid, row_keywords)

    rows = pending_rows(ws, row_keywords)
    if args.max_keywords:
        rows = rows[: args.max_keywords]

    chunk = math.ceil(len(rows) / args.workers) if rows else 0
    jobs = [rows[j * chunk : (j + 1) * chunk] for j in range(args.workers)]

    # 분할 무결성 검증: 합집합 == 원본, 중복 없음 (동일행 기록 원천 차단)
    flat = [r for job in jobs for r in job]
    assert sorted(flat) == rows, "워커 배분 검증 실패: 배분 결과가 원본과 불일치"
    assert len(flat) == len(set(flat)), "워커 배분 검증 실패: 행 중복 발생"

    config = {"tab": tab, "jobs": jobs}
    Path(args.prepare_out).write_text(json.dumps(config), encoding="utf-8")

    sizes = [len(j) for j in jobs]
    print(f"탭: {tab} | 미완료 {len(rows)}행 → 워커 {args.workers}개 분할 {sizes} → {args.prepare_out}")


# ── summary (노출 합계) ───────────────────────────────────────────

def run_summary(args) -> None:
    """날짜 탭의 순위 컬럼에서 클립 노출 키워드 수를 집계해 마지막 행에 기록.

    노출 판정: 순위(B열) 값이 숫자로 시작 ("3", "3,5" → 노출,
    "X"/"섹션없음"/공란 → 제외). 재실행 시 같은 행에 멱등하게 덮어쓴다.
    """
    sid = _spreadsheet_id(args.sheets_id)
    row_keywords, _ = load_source(sid)
    ws, tab = ensure_result_tab(sid, row_keywords, tab_name=args.tab)

    vals = _api_call(ws.get_all_values)

    last_idx = len(vals)
    if last_idx > 0 and vals[-1] and vals[-1][0] in ("합계", "합계 노출"):
        summary_row_idx = last_idx  # 기존 합계 행 덮어쓰기 (재실행 멱등)
        keyword_rows = vals[1:-1]   # 합계 행은 집계에서 제외 (자기 자신 카운트 방지)
    else:
        summary_row_idx = last_idx + 1
        if summary_row_idx > ws.row_count:
            _api_call(ws.resize, rows=summary_row_idx)
        keyword_rows = vals[1:]

    exposed = sum(
        1
        for row in keyword_rows
        if len(row) > 1 and re.match(r"^\d", str(row[1]).strip())
    )

    _api_call(ws.update, [["합계 노출", exposed]], f"A{summary_row_idx}",
              value_input_option="USER_ENTERED")
    print(f"'{tab}' {summary_row_idx}행 합계 기록: 클립 노출 키워드 {exposed}개")


# ── CLI ───────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="네이버 클립 섹션 순위 검출")
    parser.add_argument("--sheets-id", help="스프레드시트 ID (기본: GOOGLE_SHEET_ID/SPREADSHEET_ID)")
    parser.add_argument("--row-indices", help="처리할 행 번호 (쉼표/범위, 예: 2,5,10-20)")
    parser.add_argument("--start", type=int, default=0, help="시작 오프셋 (--row-indices 없을 때)")
    parser.add_argument("--count", type=int, default=None, help="처리 행 수 (--row-indices 없을 때)")
    parser.add_argument("--tab", help="결과 탭명 (기본: 오늘 KST 클립순위_MMDD)")
    parser.add_argument("--headless", action="store_true", help="브라우저 숨김 모드")
    parser.add_argument("--prepare", action="store_true", help="워커 배분 config.json 생성 후 종료")
    parser.add_argument("--prepare-out", default="config.json", help="prepare 출력 파일 경로")
    parser.add_argument("--workers", type=int, default=10, help="prepare 워커 수")
    parser.add_argument("--max-keywords", type=int, default=None, help="prepare 최대 행 수 (테스트용)")
    parser.add_argument("--summary", action="store_true", help="탭 마지막 행에 노출 키워드 합계 기록 후 종료")
    args = parser.parse_args()

    if args.prepare:
        run_prepare(args)
    elif args.summary:
        run_summary(args)
    else:
        asyncio.run(run_worker(args))


if __name__ == "__main__":
    main()

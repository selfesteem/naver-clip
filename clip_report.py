#!/usr/bin/env python3
"""
네이버 클립 순위 탭 통계 (clip_report 의 데이터 계층)

날짜별 탭(클립순위_MMDD) → 일별 통계(TabStats):
  - 총 키워드 / 검색 완료 / 섹션없음 / 오류 / 미처리
  - 구간별(1-3/1-5/1-10위) 중복포함(매치 항목 수) · 중복제거(키워드 수, 고유 영상 수)
  - 노출 영상: (매칭채널, 클립제목) 쌍 기준 고유

보고서 텍스트 조립은 clip_report_fmt, CLI 는 이 파일의 main().
환경변수는 clip_rank 와 동일 (GOOGLE_CREDENTIALS / GOOGLE_SHEET_ID ...).
"""

import argparse
import os
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Final

from clip_rank import (
    KST,
    RESULT_TAB_PREFIX,
    _env_credentials,
    _spreadsheet_id,
    result_tab_name,
)
from sheets_io import _api_call, _get_client

TIERS: Final[tuple[int, ...]] = (3, 5, 10)          # 구간 전체
COMPARE_TIERS: Final[tuple[int, ...]] = (5, 10)     # 전주/평균 대비 구간
LOOKBACK_DAYS: Final[int] = 45                      # 히스토리 탭 로드 범위
SUMMARY_ROWS: Final[frozenset[str]] = frozenset({"합계", "합계 노출"})

_TAB_RE = re.compile(rf"^{RESULT_TAB_PREFIX}(\d{{4}})$")


# ── 값 객체 ────────────────────────────────────────────────────────

@dataclass(frozen=True, slots=True)
class Match:
    rank: int
    channel: str
    title: str


@dataclass(frozen=True, slots=True)
class TierStats:
    entries: int      # 중복포함: 구간 내 매치 항목 수
    keywords: int     # 중복제거: 구간 내 노출 키워드 수
    videos: int       # 중복제거: 구간 내 고유 영상 수


@dataclass(frozen=True, slots=True)
class TabStats:
    day: date
    total: int        # 총 키워드 행
    processed: int    # 검색 완료 (노출+미노출+섹션없음)
    section_none: int
    missing: int      # 미처리 (순위 비어 있고 오류 없음)
    error: int
    exposed_any: int      # 전 구간 노출 키워드 수
    videos_any: int       # 전 구간 고유 영상 수
    tiers: dict[int, TierStats]

    def tier_rate(self, n: int) -> float:
        """구간 노출률(%) = 노출 키워드 / 검색 완료."""
        return self.tiers[n].keywords * 100 / self.processed if self.processed else 0.0


# ── 탭 파싱/통계 ──────────────────────────────────────────────────

def _split_matches(row: list[str]) -> list[Match]:
    """결과 행 → 매치 목록. 순위·채널·제목은 동일 순서로 join 되어 있다."""
    rank_val = (row[1] if len(row) > 1 else "").strip()
    if not rank_val[:1].isdigit():
        return []
    ranks = [int(p) for p in rank_val.split(",") if p.strip().isdigit()]
    channels = [c.strip() for c in (row[2] if len(row) > 2 else "").split(",")]
    titles = [t.strip() for t in (row[4] if len(row) > 4 else "").split(" / ")]
    return [
        Match(ranks[i],
              channels[i] if i < len(channels) else "",
              titles[i] if i < len(titles) else "")
        for i in range(len(ranks))
    ]


def tab_stats(day: date, vals: list[list[str]]) -> TabStats:
    """탭 전체 값 → 일별 통계."""
    total = processed = section_none = missing = error = 0
    kw_any: set[str] = set()
    vid_any: set[tuple[str, str]] = set()
    tier_kw: dict[int, set[str]] = {n: set() for n in TIERS}
    tier_vid: dict[int, set[tuple[str, str]]] = {n: set() for n in TIERS}
    tier_entries: dict[int, int] = {n: 0 for n in TIERS}

    for row in vals[1:]:  # index 0 = 헤더
        kw = (row[0] if row else "").strip()
        if not kw or kw in SUMMARY_ROWS:
            continue
        total += 1
        rank_val = (row[1] if len(row) > 1 else "").strip()
        err = (row[7] if len(row) > 7 else "").strip()
        if not rank_val:
            if err:
                error += 1
            else:
                missing += 1
            continue
        processed += 1
        if rank_val == "섹션없음":
            section_none += 1
            continue
        if rank_val == "X":
            continue
        matches = _split_matches(row)
        if matches:
            kw_any.add(kw)
        for m in matches:
            vid_any.add((m.channel, m.title))
            for n in TIERS:
                if m.rank <= n:
                    tier_entries[n] += 1
                    tier_kw[n].add(kw)
                    tier_vid[n].add((m.channel, m.title))

    tiers = {
        n: TierStats(tier_entries[n], len(tier_kw[n]), len(tier_vid[n]))
        for n in TIERS
    }
    return TabStats(day, total, processed, section_none, missing, error,
                    len(kw_any), len(vid_any), tiers)


def parse_tab_date(title: str, today: date) -> date | None:
    """'클립순위_0921' → 날짜. 연도 추정: 미래면 전년도."""
    m = _TAB_RE.match(title)
    if not m:
        return None
    mmdd = m.group(1)
    for year in (today.year, today.year - 1):
        try:
            d = date(year, int(mmdd[:2]), int(mmdd[2:]))
        except ValueError:
            continue
        if d <= today:
            return d
    return None


# ── 시트 로드 ─────────────────────────────────────────────────────

def load_history(spreadsheet_id: str, today: date) -> dict[date, TabStats]:
    """lookback 내 클립순위 탭 전체 → {날짜: 통계}."""
    _env_credentials()
    client = _get_client()
    ss = client.open_by_key(spreadsheet_id)
    cutoff = today - timedelta(days=LOOKBACK_DAYS)
    history: dict[date, TabStats] = {}
    for w in _api_call(ss.worksheets):
        day = parse_tab_date(w.title, today)
        if day is None or day < cutoff:
            continue
        history[day] = tab_stats(day, _api_call(w.get_all_values))
    return history


# ── CLI ───────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="네이버 클립 일별 노출 리포트")
    parser.add_argument("--sheets-id", help="스프레드시트 ID (기본: GOOGLE_SHEET_ID/SPREADSHEET_ID)")
    parser.add_argument("--tab", help="기준 탭명 (기본: 오늘 KST 클립순위_MMDD)")
    args = parser.parse_args()

    tab = args.tab or result_tab_name()
    today = parse_tab_date(tab, datetime.now(KST).date())
    if today is None:
        print(f"탭명을 클립순위_MMDD 형식으로 지정하세요: {tab!r}", file=sys.stderr)
        return 2

    history = load_history(_spreadsheet_id(args.sheets_id), today)
    if today not in history:
        print(f"탭을 찾을 수 없습니다: {tab!r}", file=sys.stderr)
        return 2

    # 늦은 import: clip_report_fmt 가 이 모듈의 타입을 참조 (순환 방지)
    from clip_report_fmt import build_report, gemini_trend, send_chat_report

    report = build_report(history, today)

    insight = gemini_trend(report)
    if insight:
        report = report.rstrip() + "\n\nAI 추세 요약\n" + insight + "\n"
        print("Gemini 추세 요약 반영됨")

    print(report)

    summary_path = os.environ.get("GITHUB_STEP_SUMMARY", "")
    if summary_path:
        with open(summary_path, "a", encoding="utf-8") as f:
            f.write(f"\n```\n{report}```\n")

    webhook = os.environ.get("GOOGLE_CHAT_WEBHOOK", "")
    if webhook:
        sent = send_chat_report(report, webhook)
        print(f"Google Chat 전송 완료: {sent}개 메시지")
    return 0


if __name__ == "__main__":
    sys.exit(main())

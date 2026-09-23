#!/usr/bin/env python3
"""
clip_report 단위 검증 — 합성 탭 데이터로 통계·파싱·보고서 검증 (시트 접근 없음)

사용법:
    python clip_report_test.py    # 전체 통과 시 exit 0, 실패 시 exit 1
"""

import json
import sys
from datetime import date, timedelta

import clip_report_fmt
from clip_report import TabStats, parse_tab_date, tab_stats
from clip_report_fmt import (
    _delta_int,
    _delta_pct,
    _split_chunks,
    build_report,
    send_chat_report,
)

HEADER = ["키워드", "순위", "매칭채널", "매칭조건", "클립제목", "총클립수", "처리완료", "오류"]


# ── 픽스처 ────────────────────────────────────────────────────────

def sample_tab() -> list[list[str]]:
    """kw1: 1/2/4위 3매치, kw2: 2위 (kw1과 같은 영상), X/섹션없음/미처리/오류 1행씩."""
    return [
        HEADER,
        ["kw1", "1,2,4", "chA, chB, chA", "시트채널 | URL | 시트채널", "t1 / t2 / t1", "4", "Y", ""],
        ["kw2", "2", "chA", "시트채널", "t1", "4", "Y", ""],
        ["kw3", "X", "", "", "", "5", "Y", ""],
        ["kw4", "섹션없음", "", "", "", "", "Y", ""],
        ["kw5", "", "", "", "", "", "", ""],
        ["kw6", "", "", "", "", "", "", "Timeout"],
        ["합계 노출", "2", "", "", "", "", "", ""],
    ]


def stats_for(day: date, exposed: int = 2) -> TabStats:
    """노출 exposed개·미노출 2개인 탭 통계 (비교 섹션 검증용)."""
    rows = [HEADER]
    for i in range(exposed):
        rows.append([f"kw{i}", "1", "chA", "시트채널", "t1", "4", "Y", ""])
    rows.extend(
        [["kwx", "X", "", "", "", "5", "Y", ""], ["kwy", "X", "", "", "", "5", "Y", ""]]
    )
    return tab_stats(day, rows)


# ── 테스트 ────────────────────────────────────────────────────────

def test_tab_stats_counts() -> None:
    # Given: sample_tab / When: tab_stats / Then: 행 분류·구간 집계
    s = tab_stats(date(2026, 9, 22), sample_tab())
    assert (s.total, s.processed, s.section_none, s.missing, s.error) == (6, 4, 1, 1, 1), s
    assert (s.exposed_any, s.videos_any) == (2, 2), s  # (chA,t1),(chB,t2) 고유 영상
    assert s.tiers[3].entries == 3, s.tiers[3]          # kw1의 1·2위 + kw2의 2위
    assert s.tiers[3].keywords == 2, s.tiers[3]
    assert s.tiers[10].entries == 4, s.tiers[10]        # kw1의 4위 추가
    assert s.tier_rate(3) == 50.0, s.tier_rate(3)       # 2 / 검색완료 4


def test_parse_tab_date() -> None:
    # Given: 탭명·기준일 조합 / When: parse_tab_date / Then: 연도 추정 날짜
    assert parse_tab_date("클립순위_0921", date(2026, 9, 22)) == date(2026, 9, 21)
    assert parse_tab_date("클립순위_0102", date(2026, 1, 5)) == date(2026, 1, 2)
    assert parse_tab_date("클립순위_1228", date(2027, 1, 2)) == date(2026, 12, 28)  # 연말 롤오버
    assert parse_tab_date("키워드/채널", date(2026, 9, 22)) is None
    assert parse_tab_date("블로그순위_0921", date(2026, 9, 22)) is None


def test_delta_format() -> None:
    # Given: 증감값 / When: 포맷 / Then: 부호·개수·%p
    assert _delta_int(5, 5) == "±0개"
    assert _delta_int(5, 8) == "▲3개"
    assert _delta_pct(10.0, 9.6) == "▼0.4%p"
    assert _delta_pct(1.0, 1.0) == "±0.0%p"


def test_build_report_with_history() -> None:
    # Given: 오늘·전일·전주(상승) 히스토리 / When: build_report / Then: 핵심 섹션·증감 표기
    today = date(2026, 9, 22)
    history = {
        today: stats_for(today, exposed=3),
        today - timedelta(days=1): stats_for(today - timedelta(days=1), exposed=2),
        today - timedelta(days=7): stats_for(today - timedelta(days=7), exposed=1),
        today - timedelta(days=8): stats_for(today - timedelta(days=8), exposed=1),
    }
    r = build_report(history, today)
    assert "총 키워드 5개" in r, r
    assert "데이터: 09/21 → 09/22" in r, r
    for section in ("중복포함", "중복제거 (키워드 단위)", "전주 동일(09/15)",
                    "9월 평균 대비", "주 평균 대비"):
        assert section in r, f"섹션 누락: {section}\n{r}"
    for tier in ("1-3위", "1-5위", "1-10위"):
        assert tier in r, f"구간 누락: {tier}\n{r}"
    assert "▲1개" in r and "▲10.0%p" in r, r  # 전일 2→3개, 50.0%→60.0%


def test_build_report_solo_day() -> None:
    # Given: 당일 탭만 존재 / When: build_report / Then: 비교 섹션 우회 처리
    today = date(2026, 9, 22)
    r = build_report({today: stats_for(today)}, today)
    assert "총 키워드 4개" in r, r  # 노출 2 + 미노출 2
    assert r.count("비교 데이터 없음") >= 4, r  # 전일/전주/월평균/주평균
    assert "비교 기준 데이터가 없습니다." in r, r


def test_build_report_missing_today() -> None:
    # Given: 당일 통계 없음 / When: build_report / Then: LookupError
    try:
        build_report({}, date(2026, 9, 22))
    except LookupError:
        return
    raise AssertionError("LookupError 미발생")


def test_split_chunks() -> None:
    # Given: 행 단위 텍스트 / When: limit 이하 분할 / Then: 행 경계 분할·무손실
    text = "".join(f"행{i:03d} 가나다라마바사\n" for i in range(100))
    chunks = _split_chunks(text, limit=60)
    assert all(len(c) <= 60 for c in chunks), chunks
    assert "".join(chunks) == text, chunks
    assert len(chunks) > 1, chunks
    assert _split_chunks("짧은 텍스트\n") == ["짧은 텍스트\n"]


def test_send_chat_report_payload() -> None:
    # Given: 가짜 urlopen / When: send_chat_report / Then: 청크별 JSON text 페이로드
    calls: list[bytes] = []

    class FakeResp:
        def __enter__(self) -> "FakeResp":
            return self

        def __exit__(self, *args: object) -> None:
            return None

    orig = clip_report_fmt.urllib.request.urlopen

    def fake_urlopen(req: object, timeout: int = 0) -> FakeResp:
        calls.append(getattr(req, "data"))
        return FakeResp()

    clip_report_fmt.urllib.request.urlopen = fake_urlopen
    try:
        sent = send_chat_report("가" * 500 + "\n" * 2 + "나\n", "https://example.invalid/hook")
    finally:
        clip_report_fmt.urllib.request.urlopen = orig
    assert sent == len(calls) >= 1, (sent, len(calls))
    for data in calls:
        payload = json.loads(data)
        assert set(payload) == {"text"} and payload["text"], payload


# ── 실행 ──────────────────────────────────────────────────────────

def main() -> int:
    tests = [
        test_tab_stats_counts,
        test_parse_tab_date,
        test_delta_format,
        test_build_report_with_history,
        test_build_report_solo_day,
        test_build_report_missing_today,
        test_split_chunks,
        test_send_chat_report_payload,
    ]
    failed = 0
    for t in tests:
        try:
            t()
            print(f"PASS {t.__name__}")
        except AssertionError as e:
            failed += 1
            print(f"FAIL {t.__name__}: {e}")
    print(f"\n{len(tests) - failed}/{len(tests)} 통과")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())

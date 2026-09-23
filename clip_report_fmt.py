#!/usr/bin/env python3
"""네이버 클립 일별 리포트 — 텍스트 조립·전송 (clip_report 의 표현 계층)"""

import json
import os
import sys
import urllib.request
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Final

from clip_report import COMPARE_TIERS, TIERS, TabStats

CHAT_CHUNK_LIMIT: Final[int] = 4000  # Google Chat 메시지 한도 4,096자 - 여유분


# ── 값 객체 ────────────────────────────────────────────────────────

@dataclass(frozen=True, slots=True)
class TierAvg:
    count: float
    rate: float


@dataclass(frozen=True, slots=True)
class Baseline:
    name: str
    day: date
    stats: TabStats


# ── 포맷 헬퍼 ─────────────────────────────────────────────────────

def _md(d: date) -> str:
    return f"{d.month:02d}/{d.day:02d}"


def _cnt(n: int) -> str:
    return f"{n:,}"


def _pct(x: float) -> str:
    return f"{x:.1f}%"


def _mark(before: float, after: float) -> str:
    return "▲" if after > before else "▼" if after < before else "±"


def _delta_int(before: int, after: int) -> str:
    return f"{_mark(before, after)}{abs(after - before)}개"


def _delta_pct(before: float, after: float) -> str:
    return f"{_mark(before, after)}{abs(after - before):.1f}%p"


def _tier_label(n: int) -> str:
    return f"1-{n}위"


def _line_dup_incl(cur: TabStats, prev: TabStats, n: int) -> str:
    p, c = prev.tiers[n].entries, cur.tiers[n].entries
    return f"● {_tier_label(n)}: {_cnt(p)}개 ➔ {_cnt(c)}개 ({_delta_int(p, c)})"


def _line_dedup(cur: TabStats, prev: TabStats, n: int) -> str:
    p, c = prev.tiers[n].keywords, cur.tiers[n].keywords
    return (f"● {_tier_label(n)}: {_cnt(p)}개 / {_pct(prev.tier_rate(n))} ➔ "
            f"{_cnt(c)}개 / {_pct(cur.tier_rate(n))} "
            f"({_delta_int(p, c)} / {_delta_pct(prev.tier_rate(n), cur.tier_rate(n))})")


def _line_vs_avg(cur: TabStats, avg: TierAvg, n: int) -> str:
    """월/주 평균 대비 한 줄: '평균값' [기준일 대비 증감]."""
    avg_cnt = round(avg.count)
    return (f" · {_tier_label(n)} : {_cnt(avg_cnt)}개 / {avg.rate:.1f}% "
            f"[{_delta_int(avg_cnt, cur.tiers[n].keywords)} / "
            f"{_delta_pct(avg.rate, cur.tier_rate(n))}]")


def _mean(values: list[float]) -> float:
    return sum(values) / len(values) if values else 0.0


def _tier_avg(days: list[TabStats], n: int) -> TierAvg:
    return TierAvg(
        _mean([float(d.tiers[n].keywords) for d in days]),
        _mean([d.tier_rate(n) for d in days]),
    )


# ── 요약 문장 ─────────────────────────────────────────────────────

def _trend_phrase(diff: int, pp: float) -> str:
    """개수·노출률 증감을 종합한 트렌드 문구. 방향이 어긋나면 검색 완료 수 변동 명시."""
    if diff == 0 and abs(pp) < 0.05:
        return "대체로 유지"
    cnt_up, rate_up = diff > 0, pp > 0
    if cnt_up != rate_up and diff != 0 and abs(pp) >= 0.05:
        return (f"노출 수 {'상승' if cnt_up else '하락'}, "
                f"노출률 {'상승' if rate_up else '하락'} (검색 완료 수 변동)")
    mag = "급격한" if abs(pp) >= 3.0 else "중폭" if abs(pp) >= 1.0 else "소폭"
    return f"{mag} {'상승' if cnt_up or rate_up else '하락'}"


def _narrative(cur: TabStats, prev: Baseline | None,
               week: Baseline | None) -> list[str]:
    lines: list[str] = []
    for base in (prev, week):
        if base is None:
            continue
        n = COMPARE_TIERS[0]
        diff = cur.tiers[n].keywords - base.stats.tiers[n].keywords
        pp = cur.tier_rate(n) - base.stats.tier_rate(n)
        lines.append(
            f"{base.name}({_md(base.day)}) 1-5위 중복제거 {_delta_int(0, diff)}"
            f"({_delta_pct(0, pp)}) — {_trend_phrase(diff, pp)}."
        )
    issues: list[str] = []
    if cur.error:
        issues.append(f"오류 {cur.error}행")
    if cur.missing:
        issues.append(f"미처리 {_cnt(cur.missing)}행 — 노출률은 검색 완료({_cnt(cur.processed)}행) 기준")
    if issues:
        lines.append("※ " + " · ".join(issues))
    if not lines:
        lines.append("비교 기준 데이터가 없습니다.")
    return lines


# ── 보고서 생성 ───────────────────────────────────────────────────

def build_report(history: dict[date, TabStats], today: date) -> str:
    """히스토리 통계 → 보고서 텍스트 (순수 함수)."""
    cur = history.get(today)
    if cur is None:
        raise LookupError(f"{today} 탭 통계가 없습니다.")

    prev_days = [d for d in history if d < today]
    prev_day = max(prev_days) if prev_days else None
    prev = Baseline("전일 대비", prev_day, history[prev_day]) if prev_day else None
    week_day = today - timedelta(days=7)
    week = (Baseline("전주 동일 요일", week_day, history[week_day])
            if week_day in history else None)

    out: list[str] = [
        f"● 네이버 클립 노출 ({today} 기준)",
        "",
        f"총 키워드 {_cnt(cur.total)}개 (검색 완료 {_cnt(cur.processed)} · "
        f"섹션없음 {_cnt(cur.section_none)} · 오류 {_cnt(cur.error)} · 미처리 {_cnt(cur.missing)})",
        f"노출 영상 {_cnt(cur.videos_any)}개 · 영상 노출 키워드 {_cnt(cur.exposed_any)}개 · "
        f"중복제거 노출률 {_pct(cur.exposed_any * 100 / cur.processed if cur.processed else 0.0)}",
        "",
    ]
    out.extend(_narrative(cur, prev, week))
    out.append("")

    # 전일 대비 상세
    out.append(f"데이터: {_md(prev.day)} → {_md(today)}" if prev
               else "데이터: 비교 가능한 이전 탭 없음")
    out.append("")
    out.append("중복포함 (네이버클립 노출 기준)")
    if prev:
        out.extend(_line_dup_incl(cur, prev.stats, n) for n in TIERS)
    else:
        out.extend(f"● {_tier_label(n)}: {_cnt(cur.tiers[n].entries)}개 (비교 데이터 없음)"
                   for n in TIERS)
    out.append("")
    out.append("중복제거 (네이버클립 노출 기준 - 키워드 단위)")
    if prev:
        out.extend(_line_dedup(cur, prev.stats, n) for n in TIERS)
    else:
        out.extend(f"● {_tier_label(n)}: {_cnt(cur.tiers[n].keywords)}개 / {_pct(cur.tier_rate(n))} (비교 데이터 없음)"
                   for n in TIERS)
    out.append("")

    # 전주 동일 요일 대비
    out.append(f"전주 동일({_md(week_day)}) 대비 중복제거")
    if week:
        out.extend(
            f"● {_tier_label(n)}: {_cnt(week.stats.tiers[n].keywords)}개"
            f"({week.stats.tier_rate(n):.1f}%) ➔ {_cnt(cur.tiers[n].keywords)}개"
            f"({cur.tier_rate(n):.1f}%) "
            f"({_delta_int(week.stats.tiers[n].keywords, cur.tiers[n].keywords)} / "
            f"{_delta_pct(week.stats.tier_rate(n), cur.tier_rate(n))})"
            for n in COMPARE_TIERS
        )
    else:
        out.append("· 비교 데이터 없음")
    out.append("")

    # 당월 평균 대비
    month_days = [s for d, s in sorted(history.items())
                  if d.month == today.month and d != today]
    out.append(f"{today.month}월 평균 대비 중복제거 [기준일 대비 증감]")
    if month_days:
        out.extend(_line_vs_avg(cur, _tier_avg(month_days, n), n)
                   for n in COMPARE_TIERS)
    else:
        out.append("· 비교 데이터 없음")
    out.append("")

    # 주 평균 대비 (완결된 주 최대 2주)
    week_starts = sorted({d - timedelta(days=d.weekday()) for d in history})
    completed = [ws for ws in week_starts if ws + timedelta(days=6) < today][-2:]
    out.append("주 평균 대비 중복제거")
    shown = 0
    for ws in completed:
        days = [s for d, s in sorted(history.items())
                if ws <= d <= ws + timedelta(days=6)]
        if not days:
            continue
        out.append(f"{ws} - {ws + timedelta(days=6)}")
        out.extend(_line_vs_avg(cur, _tier_avg(days, n), n)
                   for n in COMPARE_TIERS)
        shown += 1
    if not shown:
        out.append("· 비교 데이터 없음")
    return "\n".join(out).rstrip() + "\n"


# ── Gemini 추세 요약 (선택) ────────────────────────────────────────

GEMINI_MODEL_DEFAULT: Final[str] = "gemini-2.5-flash"

_TREND_PROMPT = """아래 네이버 클립 노출 일별 리포트를 읽고 추세 요약 2줄을 한국어로 작성하세요.
규칙:
- 첫째 줄: 전일·전주·평균 대비 노출 수/노출률 변화를 리포트에 있는 수치로만 요약
- 둘째 줄: 리포트에서 확인되는 특징(구간별 편차, 오류/미처리 등)에 근거한 점검 또는 활용 방향 1문장
- 리포트에 없는 수치·사실 추측 금지, 각 줄 100자 이내, 인사말·마크다운·이모지 없이 본문만

리포트:
"""


def _extract_trend_text(data: dict) -> str:
    candidates = data.get("candidates") or []
    parts = (candidates[0].get("content") or {}).get("parts") or [] if candidates else []
    return "".join(p.get("text") or "" for p in parts if not p.get("thought")).strip()


def gemini_trend(report: str) -> str:
    """리포트 통계 → Gemini 2줄 추세 요약.

    GEMINI_API_KEY 가 없으면 기능 꺼짐(빈 값). 호출 실패 시에도 빈 값을
    반환해 리포트 본문이 영향받지 않는다. 모델은 GEMINI_MODEL 로
    오버라이드 가능 (기본 gemini-2.5-flash).
    thinkingBudget 0: 2.5계열 추론 토큰이 maxOutputTokens 예산을 먼저
    소진해 답변이 잘리는 현상(유료 토큰 낭비 포함)을 막는다.
    """
    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        return ""
    model = os.environ.get("GEMINI_MODEL", "").strip() or GEMINI_MODEL_DEFAULT
    payload = json.dumps({
        "contents": [{"parts": [{"text": _TREND_PROMPT + report}]}],
        "generationConfig": {
            "temperature": 0.3,
            "maxOutputTokens": 1024,
            "thinkingConfig": {"thinkingBudget": 0},
        },
    }).encode()
    req = urllib.request.Request(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        data=payload,
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode())
    except Exception as e:
        print(f"Gemini 추세 요약 실패(스킵): {e}", file=sys.stderr)
        return ""
    return _extract_trend_text(data)


# ── Google Chat 전송 ──────────────────────────────────────────────

def _split_chunks(text: str, limit: int = CHAT_CHUNK_LIMIT) -> list[str]:
    """텍스트를 행 경계 기준 limit 이하 청크로 분할 (내용 무손실)."""
    chunks: list[str] = []
    buf = ""
    for line in text.splitlines(keepends=True):
        if buf and len(buf) + len(line) > limit:
            chunks.append(buf)
            buf = ""
        buf += line
    if buf:
        chunks.append(buf)
    return chunks


def send_chat_report(text: str, webhook_url: str) -> int:
    """리포트를 Google Chat 웹훅으로 전송. 전송한 메시지 수 반환."""
    sent = 0
    for chunk in _split_chunks(text):
        payload = json.dumps({"text": chunk}).encode()
        req = urllib.request.Request(
            webhook_url,
            data=payload,
            headers={"Content-Type": "application/json; charset=utf-8"},
        )
        with urllib.request.urlopen(req, timeout=15):
            sent += 1
    return sent

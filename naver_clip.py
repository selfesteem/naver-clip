"""
네이버 모바일 검색결과 전체 섹션 순위 확인 모듈
"""

import asyncio
import os
from urllib.parse import quote
from playwright.async_api import async_playwright, Page

MOBILE_UA = (
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1"
)
MOBILE_VIEWPORT = {"width": 390, "height": 844}

# 섹션명 정규화 (모바일 페이지 표기 → 표준명)
SECTION_ALIASES: dict[str, str] = {
    "클립": "네이버 클립",
    "CLIP": "네이버 클립",
    "블로그 글": "블로그",
    "인플루언서 글": "인플루언서",
    "카페글": "카페",
    "지식인": "지식iN",
    "네이버 지식iN": "지식iN",
}

_SECTION_JS = """
(keywords) => {
    function hasTarget(html) {
        return keywords.some(k => html.includes(k));
    }
    // script/style 제거 후 innerHTML 검사 (JSON 오탐 방지 + img alt 등 속성값 포함)
    function hasVisibleTarget(el) {
        const clone = el.cloneNode(true);
        clone.querySelectorAll('script, style').forEach(e => e.remove());
        return keywords.some(k => clone.innerHTML.includes(k));
    }

    const results = [];

    // 1. h2 레이블이 있는 named 섹션
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
            results.push({ name, has_target: hasVisibleTarget(sec), position: null, total: 0 });
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

        let position = null;
        for (let i = 0; i < items.length; i++) {
            if (hasTarget(items[i].innerHTML || '')) { position = i + 1; break; }
        }

        results.push({
            name,
            has_target: position !== null || (items.length === 0 && hasVisibleTarget(sec)),
            position,
            total: items.length,
        });
    }

    // 2. spw_fsolid 섹션들 (섹션명 없는 영역 → 웹문서 1, 2, ...)
    const fsolidSections = Array.from(document.querySelectorAll('.spw_fsolid'));
    fsolidSections.forEach((sec, idx) => {
        const list = sec.querySelector('.fsolid_list');
        const items = list
            ? Array.from(list.children).filter(el => el.tagName === 'DIV')
            : Array.from(sec.children).filter(el => el.tagName === 'DIV');
        if (items.length === 0) return;

        let position = null;
        for (let i = 0; i < items.length; i++) {
            if (hasTarget(items[i].innerHTML || '')) { position = i + 1; break; }
        }
        const name = `웹문서 ${idx + 1}`;
        results.push({
            name,
            has_target: position !== null,
            position,
            total: items.length,
        });
    });

    // 3. 플레이스 영역
    const placeSection = document.querySelector('[class*="place_section"]');
    if (placeSection) {
        // ul.children으로 직접 자식 li만 (querySelectorAll은 중첩 li까지 잡아 위치 오산)
        let items = [];
        const ul = placeSection.querySelector('ul');
        if (ul) {
            items = Array.from(ul.children).filter(el => el.tagName === 'LI');
        }
        // ul 없으면 클래스 기반 직접 자식 div
        if (items.length === 0) {
            const container = placeSection.querySelector('[class*="place_list"], [class*="place_lst"]');
            if (container) items = Array.from(container.children);
        }
        if (items.length === 0) {
            items = Array.from(placeSection.querySelectorAll('[class*="place_item"], [class*="UEzoS"]'));
        }

        let position = null;
        for (let i = 0; i < items.length; i++) {
            if (hasTarget(items[i].textContent || '')) { position = i + 1; break; }
        }
        const has_target = position !== null || (items.length === 0 && hasTarget(placeSection.textContent || ''));
        results.push({
            name: '플레이스',
            has_target,
            position,
            total: items.length,
        });
    }

    return results;
}
"""


def _get_target_keywords() -> list[str]:
    val = os.environ.get("TARGET_KEYWORDS", "")
    return [k.strip() for k in val.split(",") if k.strip()]


def normalize_section(name: str) -> str:
    return SECTION_ALIASES.get(name, name)


async def search_all_sections(page: Page, keyword: str) -> dict:
    """
    모바일 네이버에서 keyword 검색 → 전체 섹션별 노출 현황 반환.

    반환 dict:
        keyword   : 검색어
        sections  : [{ name, has_target, position, total }, ...]
        error     : 오류 메시지 (str|None)
    """
    result = {"keyword": keyword, "sections": [], "error": None}
    try:
        url = f"https://m.search.naver.com/search.naver?query={quote(keyword)}"
        await page.goto(url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(2_000)

        raw = await page.evaluate(_SECTION_JS, _get_target_keywords())
        # 섹션명 정규화
        for sec in raw:
            sec["name"] = normalize_section(sec["name"])
        result["sections"] = raw

    except Exception as e:
        result["error"] = str(e)

    return result


async def run_inspection(keyword: str):
    """개발용: 키워드의 모바일 검색결과 섹션 구조를 출력."""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        ctx = await browser.new_context(
            user_agent=MOBILE_UA, viewport=MOBILE_VIEWPORT, locale="ko-KR"
        )
        await ctx.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        page = await ctx.new_page()
        url = f"https://m.search.naver.com/search.naver?query={quote(keyword)}"
        print(f"[inspect] {url}\n")
        await page.goto(url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(2_000)

        raw = await page.evaluate(_SECTION_JS, _get_target_keywords())
        for sec in raw:
            sec["name"] = normalize_section(sec["name"])

        if not raw:
            print("섹션을 찾지 못했습니다. 페이지 구조 디버그:")
            debug = await page.evaluate("""() =>
                Array.from(document.querySelectorAll('section, [class*="sc_"], [class*="area"]'))
                    .slice(0, 25)
                    .map(el => ({
                        tag: el.tagName,
                        cls: el.className.slice(0, 70),
                        h2: ((el.querySelector('h2') || {}).textContent || '').trim().slice(0, 40),
                    }))
            """)
            for d in debug:
                print(f"  <{d['tag']} class=\"{d['cls']}\">  h2='{d['h2']}'")
        else:
            # 플레이스 섹션 구조 추가 디버그
            place_debug = await page.evaluate("""() => {
                const sec = document.querySelector('[class*="place_section"]');
                if (!sec) return null;
                const ul = sec.querySelector('ul');
                if (!ul) return { cls: sec.className, ul: null };
                const lis = Array.from(ul.children).filter(el => el.tagName === 'LI');
                return {
                    cls: sec.className.slice(0, 80),
                    ul_cls: ul.className.slice(0, 80),
                    li_count: lis.length,
                    items: lis.map((li, i) => ({
                        idx: i,
                        cls: li.className.slice(0, 60),
                        text: li.textContent.trim().slice(0, 50),
                    })),
                };
            }""")
            if place_debug:
                print(f"\n[플레이스 구조] sec.class={place_debug.get('cls')}")
                print(f"  ul.class={place_debug.get('ul_cls')}  li수={place_debug.get('li_count')}")
                for item in (place_debug.get('items') or []):
                    print(f"  [{item['idx']+1}] {item['cls']} | {item['text']}")
            print(f"섹션 {len(raw)}개 발견:\n")
            for sec in raw:
                if sec["has_target"] and sec["position"]:
                    mark = f"  ★ {sec['position']}위 / {sec['total']}개 중"
                elif sec["has_target"]:
                    mark = "  ★ 있음 (위치 미확인)"
                else:
                    mark = ""
                print(f"  [{sec['name']}] 항목 {sec['total']}개{mark}")

        await browser.close()

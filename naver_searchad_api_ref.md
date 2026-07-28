# 네이버 검색광고 API 레퍼런스

출처: https://naver.github.io/searchad-apidoc/

---

## 기본 정보

| 항목 | 값 |
|---|---|
| Service URL | `https://api.searchad.naver.com` |
| API 키 발급 | searchad.naver.com → 도구 → API 사용 관리 |

---

## 인증 (공통 헤더)

모든 요청에 아래 4개 헤더 필수.

```
X-Timestamp : Unix Epoch 밀리초 (예: 1457082455307)
X-API-KEY   : 발급받은 API 라이선스 키
X-Customer  : 고객 ID (숫자)
X-Signature : 서명값 (아래 참고)
```

### 서명 생성

```
sha256-hmac( SECRET_KEY, "{timestamp}.{METHOD}.{request_uri}" )
→ Base64 인코딩

예) 1457082455307.GET./keywordstool
```

- `request_uri`는 쿼리스트링 **제외** 순수 경로만 사용 (e.g. `/keywordstool`)
- Python 구현:
  ```python
  import hmac, hashlib, base64
  msg = f"{timestamp}.GET./keywordstool"
  sig = hmac.new(secret_key.encode(), msg.encode(), hashlib.sha256).digest()
  signature = base64.b64encode(sig).decode()
  ```

---

## GET /keywordstool — 키워드 검색량 조회

### 요청 파라미터 (Query String)

| 파라미터 | 타입 | 설명 |
|---|---|---|
| `hintKeywords` | string | **쉼표 구분**, 최대 5개 (예: `이혼소송,변호사상담`) — 쉼표를 URL 인코딩(`%2C`)하면 400 오류 |
| `showDetail` | integer | `1` = 상세 통계 포함 (검색수·클릭수·CTR·경쟁정도), `0` = 검색량만 |
| `siteId` | string | 사이트 채널 비즈니스채널 ID (선택) |
| `biztpId` | integer | 업종 ID (선택) |
| `event` | integer | 시즌 테마 ID (선택) |
| `month` | integer | 월 (예: 12, 선택) |

> 최소 1개 이상 파라미터(`siteId`, `biztpId`, `hintKeywords`, `event`) 필요.

### 응답 필드 (`keywordList` 배열 원소)

| 필드 | 타입 | 설명 |
|---|---|---|
| `relKeyword` | string | 키워드 (입력 키워드 + 연관 키워드 포함) |
| `monthlyPcQcCnt` | **string** | 최근 30일 PC 검색수. 10 미만이면 `"<10"` 반환 |
| `monthlyMobileQcCnt` | **string** | 최근 30일 모바일 검색수. 10 미만이면 `"<10"` 반환 |
| `monthlyAvePcClkCnt` | string | 최근 4주 PC 평균 클릭수 (데이터 없으면 `0`) |
| `monthlyAveMobileClkCnt` | string | 최근 4주 모바일 평균 클릭수 |
| `monthlyAvePcCtr` | string | 최근 4주 PC CTR |
| `monthlyAveMobileCtr` | string | 최근 4주 모바일 CTR |
| `plAvgDepth` | string | 최근 4주 PC 광고 평균 노출 순위 |
| `compIdx` | string | 경쟁 강도: `낮음` / `중간` / `높음` (문서엔 low/mid/high로 표기되나 실제 반환값은 한글) |

> **주의**: 검색수 필드는 integer가 아닌 **string** 타입. `"<10"` 케이스를 반드시 처리해야 함.

### 주의사항

- `hintKeywords`의 쉼표는 URL 인코딩(`%2C`) 금지 — 리터럴 `,` 사용
- `requests` 라이브러리의 `params=` 딕셔너리는 쉼표를 `%2C`로 인코딩하므로 **URL 직접 조합** 필요
- 응답에는 입력 키워드 외 **연관 키워드도 포함**되므로 `relKeyword`로 매핑 필수
- 연관 키워드 응답은 입력 키워드가 항상 포함된다는 보장 없음 — 없는 경우 결과 없음으로 처리

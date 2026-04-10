# 투자 의사결정 PWA — 통합 개선 로드맵 v5.5

> **목적**: 향후 AI 또는 개발자가 이 PWA를 개선할 때 참조하는 단일 문서.
> **현재 버전**: v5.5 (2025-04-10)
> **파일 구조**: 단일 HTML (~18,867줄). screeningData JSON 하드코딩 (~517종목).
> **제외 항목**: 데이터 파이프라인 예제 데이터 이슈, Price/Close 필드 부재(외부 크롤링으로 해결).

---

## 0. 현재 아키텍처 요약

```
index.html (단일 파일)
├── <style>       : CSS 변수 시스템 (--fs-*/--font-* 이중), 다크 테마
├── <script#1>    : xlsx.js CDN
├── HTML          : 4개 메인 탭 (stocks / myposition / market / knowledge)
│                   knowledge 탭은 <template> 지연 렌더링
├── <script#2>    : 메인 로직 (~13,600줄)
│   ├── AppStorage (IndexedDB + localStorage 폴백)
│   ├── screeningData[] (517종목 하드코딩)
│   ├── 멀티팩터 스코어링 엔진 (V/Q/G/M/R, 5개 버전 A~E)
│   ├── evaluateMarket() 매크로 판단 엔진
│   ├── 포트폴리오 분석 (환 리스크/세금 손실/리스크 패리티)
│   ├── 시나리오 시스템 (정책/경제/이벤트 3카테고리, 12+시나리오)
│   ├── 스냅샷 히스토리 & 차트
│   └── 종목 상세 모달
└── <script#3>    : 스냅샷/차트 보조 로직
```

---

## 1. P0 — 즉시 수정 필요 (Critical)

### 1.1 NaN/Infinity 전파 방지
- **문제**: JS IEEE 754 산술에서 `price / eps` (eps=0 → Infinity, eps=undefined → NaN). NaN은 전파성 — 가중 스코어 전체를 오염시킴.
- **수정 방법**:
  ```javascript
  function safeDivide(num, den, fallback = null) {
    if (!Number.isFinite(num) || !Number.isFinite(den) || den === 0) return fallback;
    return num / den;
  }
  ```
  - 모든 비율 계산(P/E, P/B, D/E, 성장률, 팩터 스코어)에 적용
  - 복합 스코어 계산 시 null/NaN 메트릭은 가중치에서 제외 후 잔여 가중치 재정규화
  - `Number.isFinite()` 사용 (글로벌 `isFinite()`는 타입 강제변환 문제)

**프롬프트 예시**:
```
index.html의 모든 나눗셈 연산을 감사하세요.
safeDivide() 유틸리티를 구현하고, 스코어링 엔진(TotalScore, ValueScore, QualityScore 등)의
모든 비율 계산에 적용하세요. 복합 스코어에서 null/NaN 메트릭은 가중치 재정규화로 처리하세요.
```

### 1.2 비동기 초기화 레이스 컨디션
- **문제**: IndexedDB `onupgradeneeded` + fetch + Worker 동시 초기화. 트랜잭션이 마이크로태스크 완료 시 자동 커밋.
- **수정 방법**: 엄격한 순차 초기화 — (1) IndexedDB 열기/마이그레이션 완료 → (2) Worker 준비 → (3) 데이터 로드
  - `db.onversionchange` 핸들러 추가 (멀티탭 시나리오)
  - IDB 트랜잭션 내에서 비-IDB async 작업 금지

**프롬프트 예시**:
```
AppStorage의 초기화 시퀀스를 감사하세요. IndexedDB open/upgrade가 완료되기 전에
다른 코드가 스토리지에 접근하는 경로가 있는지 확인하고,
엄격한 순차 초기화(IDB → Worker → Data)로 리팩토링하세요.
db.onversionchange 핸들러도 추가하세요.
```

---

## 2. P1 — 주요 개선 (Major)

### 2.1 매도 신호 임계값 레짐 인식 재보정
- **현재**: 고정 임계값 (VIX>35, HY spread>5%, CAPE>35)
- **문제**:
  - VIX>35는 공포 정점(역발상 매수 신호)에 가까움 → 백분위 순위(90th of trailing 252일) + 변화율(5일 +40%) 사용
  - CAPE>35 바이너리 트리거 부적절 → 10년 기대수익률 조정자로 사용 (점진적 배분 축소)
  - HY spread>5%에 변화율 에스컬레이터 추가 (30일 +100bps)

**프롬프트 예시**:
```
evaluateMarket() 함수의 매도/방어 신호 임계값을 검토하세요.
VIX는 절대값 대신 252일 백분위 + 5일 변화율 조합으로,
CAPE는 바이너리 트리거 대신 기대수익률 감쇠 함수로,
HY spread는 절대값 + 30일 변화율 에스컬레이터로 개선하세요.
```

### 2.2 분산 리스크 프리미엄 (VRP) 신호 추가
- VRP = VIX − 30일 실현 변동성
- VRP >8%: 과잉 공포 (역발상 매수), VRP <2%: 안일함 (방어 전환)
- 기존 VIX 수준만 사용 중 → 내재 vs 실현 변동성 스프레드 정보 미활용

### 2.3 팩터 모델 구조 업그레이드

#### 2.3.1 Fama-French RMW(수익성) / CMA(투자) 팩터 추가
- RMW: Gross Profits/Assets (Novy-Marx 2013) — 가장 일관된 양의 프리미엄 (연 3-4%)
- CMA: 저자산성장 기업 프리미엄
- **필요 데이터**: 매출총이익, 영업이익, 총자산, 전년 대비 총자산 증가율

#### 2.3.2 QMJ 4-Pillar 품질 분해
- 현재 Quality: ROE, 부채비율, 수익 안정성 등 2-3개 메트릭
- QMJ (AQR): Profitability/Growth/Safety/Payout 20+ 변수
- **필요 데이터**: 발생액 품질, Altman Z-score, 자사주매입률, 배당성장률 등

#### 2.3.3 이익수정 모멘텀 (SUE/ERR)
- 현재 모멘텀: 가격 모멘텀만 사용 (12-1개월)
- SUE: (실제 EPS − 컨센서스 EPS) / σ(예측오차) → 상위 분위 6% 초과수익
- ERR: 상향 − 하향 수정 / 전체
- 가격 모멘텀 60% + 이익 모멘텀 40% 분할 권장
- **필요 데이터**: 컨센서스 EPS, 애널리스트 추정치, 실적 발표일

#### 2.3.4 Growth→Quality 통합 (P2)
- Growth는 학술적으로 독립 팩터 미지지 → Quality 하위 구성요소로 흡수
- 5개 버전 A~E 전수 수정 필요

#### 2.3.5 팩터 타이밍 via 밸류에이션 스프레드 (P2)
- Long/Short 간 상대 밸류에이션으로 팩터 수익률 예측 (1년+ 호라이즌)
- 90th 백분위 초과 시 해당 팩터 오버웨이트
- **필요 데이터**: 팩터별 월간 수익률 20년+

### 2.4 배분 엔진

#### 2.4.1 Black-Litterman 배분 엔진
- 시장균형 수익률(Π = δΣw_mkt) + 베이지안 뷰 업데이트
- MVO의 기대수익률 추정 민감성 제거
- 기존 팩터 스코어/모멘텀/시나리오에서 프로그래밍적으로 뷰 생성
- **필요**: Ledoit-Wolf 축소 공분산 행렬, 시가총액 가중치, δ≈2.5-3.5

**프롬프트 예시**:
```
Black-Litterman 배분 엔진을 구현하세요.
(1) Ledoit-Wolf 축소 공분산 행렬 추정
(2) 시가총액 가중치에서 균형 수익률 역산
(3) 기존 팩터 스코어와 시장 판단 결과를 뷰로 변환
(4) Idzorek 방식 신뢰도(0-100%) UI 슬라이더
결과를 현재 배분 추천 UI에 통합하세요.
```

#### 2.4.2 True ERC (Equal Risk Contribution)
- 현재 역변동성 가중 → 상관성 무시
- 진정한 ERC: 비선형 최적화로 각 자산의 한계 리스크 기여 균등화
- 컴포넌트 VaR/CVaR 분해 추가

#### 2.4.3 Cross-asset TSMOM 오버레이
- 자산군별 12M 시계열 모멘텀으로 배분 결정
- sign(r_{t-12,t}) + 변동성 스케일링, 1/3/12개월 룩백 블렌딩
- **필요 데이터**: 자산군별 12개월 수익률 시계열

### 2.5 매크로/리스크 신호 추가

| 신호 | 설명 | 데이터 소스 |
|------|------|------------|
| Credit Impulse | 신용 스톡의 2차 미분 / GDP, GDP 9-12개월 선행 | FRED Z.1 |
| SLOOS | 은행 대출 태도 서베이, C&I 강화 >38% → 매 경기침체 선행 | Fed 분기별 |
| ACM Term Premium | 기간 프리미엄 분해, 음수 → 듀레이션 언더웨이트 | NY Fed 일별 |
| VIX 텀스트럭처 | VIX/VIX3M >1.0 = 백워데이션 = 극단적 스트레스 | CBOE |
| CBOE SKEW | >130 → 꼬리 리스크 프라이싱 상승 | CBOE |
| Put-Call Ratio | 5일+ 지속 >1.0 = 극단적 공포 | CBOE/FRED |
| 인플레이션 레짐 | CPI YoY 임계값 (>4%/2-4%/<2%) → 자산배분 조정 | FRED |

### 2.6 환 헤지 프레임워크 (KRW)
- 채권 ~100% 헤지, 주식 30-50% 헤지
- FX 포워드 포인트 비용 표시
- USDKRW 장기 평균 이탈도 → 전술적 신호

### 2.7 세금 손실 수확 (Tax-loss Harvesting)
- 한국 해외주식 양도세 22%, 연간 공제 ₩250만
- Wash-sale 규정 미적용 (한국 세법)
- 알고리즘: 실현 이익 >₩250만 & 미실현 손실 >₩50만 포지션 → 수확 알림

**프롬프트 예시**:
```
내 포지션 탭에 Tax-loss Harvesting 알림 기능을 추가하세요.
조건: _sheet가 US_Stocks인 해외 포지션 중 미실현 손실 >₩500,000이고,
해당 연도 실현 이익 합계 >₩2,500,000일 때 알림 카드 표시.
한국 세법 기준: 양도세 22%, 연간 공제 ₩2,500,000, wash-sale 미적용.
```

### 2.8 가상 스크롤 (Virtual Scrolling)
- 종목 테이블 DOM 90%+ 절감
- spacer 요소 + IntersectionObserver + 스크롤 위치 관리
- `requestAnimationFrame` 배칭

### 2.9 접근성 (ARIA/WCAG)
- 데이터 테이블: `role="grid"`, `aria-sort`, `aria-rowcount`/`aria-rowindex`
- 탭: `role="tablist"/"tab"/"tabpanel"` + 화살표키 내비게이션 (부분 구현됨)
- 색상 단독 의존 금지: ▲/▼ 화살표 + ±부호 병행 (부분 구현됨)
- 동적 가격 업데이트: `aria-live="polite"`
- 색상 대비 최소: 텍스트 4.5:1, UI 컴포넌트 3:1

### 2.10 메모리 누수 감사
- addEventListener 23건, removeEventListener 0건 → 탭 전환 시 누적
- `window`/`document` 레벨 리스너 (resize 등) → 뷰 변경 시 미제거
- 클로저 내 분리된 DOM 노드 참조
- **수정**: ChartManager 클래스, `WeakRef` 요소 캐시, 리스너 정리

---

## 3. P2 — 기능 강화 (Enhancement)

### 3.1 투자 이론

| # | 항목 | 설명 |
|---|------|------|
| 1 | 배당 지속성 스코어링 | FCF 커버리지(40%), 수익 배당성향(20%), 배당 추세(15%), ROE 안정성(15%), 부채(10%) |
| 2 | Markov 레짐 스위칭 | Hamilton 1989, S&P 500 수익률 2-3상태 MRS → "위기 레짐 확률 >60%" 매도 신호 |
| 3 | Cornish-Fisher VaR | 비정규 분포 조정 VaR/CVaR, 포트폴리오 왜도/첨도 추적 |
| 4 | GPR/GPRNK 지정학 리스크 | Caldara & Iacoviello 지수, GPRNK >2σ → 한국 주식 3-5% 축소 |
| 5 | 유동성 리스크 (Amihud ILLIQ) | \|Return\| / Dollar Volume, 90th 백분위 초과 종목 경고 |
| 6 | 행동재무학 — 처분효과 감지 | PGR/PLR 비율 추적, >1.5 시 경고 |
| 7 | China Hard Landing 시나리오 | 한국 반도체 수출 25% 의존, 독립 시나리오로 추가 |
| 8 | 시나리오 확률 프레임워크 | 베이지안 사전/사후 확률 (연착륙 40-50%, 침체 15-25% 등) |
| 9 | 임계값 기반 리밸런싱 | 200bps 이탈 + 175bps 목적지 회랑 (Vanguard 2022/2024 연구) |
| 10 | 섹터 순환 가이드 | 매크로 레짐→Fidelity식 섹터 틸트 (정보 오버레이) |
| 11 | REIT/대안 배분 | 최적 REIT 5-15%, K-REITs + 글로벌 REIT ETF |

### 3.2 코드 품질

| # | 항목 | 설명 |
|---|------|------|
| 1 | `Promise.allSettled` | 데이터 로딩 부분 실패 시 graceful degradation |
| 2 | `content-visibility` + Safari | `contain-intrinsic-size: auto <fallback>` 확인, Safari Cmd+F 미검색 주의 |
| 3 | 인라인 스타일 → CSS 클래스 추출 | renderMyHoldings() 등 JS innerHTML 내 인라인 style 다수 |
| 4 | var → const/let | 블록 스코핑 전환 |
| 5 | screeningData 외부 분리 | `JSON.parse()`로 로드 (V8에서 객체 리터럴 대비 1.7× 빠름) |
| 6 | Web Worker 클로저 리스크 | `Function.toString()` 기반 인라인 Worker → 외부 스코프 변수 undefined |
| 7 | `URL.revokeObjectURL()` | Worker 생성 후 Blob URL 해제 누락 |
| 8 | 18,000줄 단일 파일 리스크 | `dynamic import()`, `scheduler.yield()`, 초기화 단계별 양보 |

### 3.3 UX/디자인

| # | 항목 | 설명 |
|---|------|------|
| 1 | 내 포지션 탭 정보 계층 | 총 가치(28-32px) → 일일 손익 → 도넛 차트 → 카드형 보유 목록 |
| 2 | 실행 가능성 (Actionability) | "모닝 브리프" 카드, 리밸런싱 알림, 포지션 신호등(🔴🟡🟢), 처분효과 경고 |
| 3 | 하단 탭 라벨 판독성 | 480px에서 0.55rem → 아이콘 전용 + 선택 시 라벨 노출 |
| 4 | 빈 상태 UI | 보유종목 0건 시 일러스트 + CTA |
| 5 | 탭 전환 시 스크롤 미복원 | 각 탭의 scrollTop 저장/복원 |
| 6 | Config JSON 에러 피드백 | alert() → 인라인 에러 + 하이라이팅 |
| 7 | 로딩 스켈레톤 | 스피너 대신 shimmer 와이어프레임, 점진적 로드 |
| 8 | 모바일 테이블 | 첫 열 고정 + 가로 스크롤 + 오른쪽 끝 잘림 힌트 |
| 9 | 디자인 시스템 카드 변형 | MetricCard, StockCard, AlertCard, InsightCard 4종 |
| 10 | 폰트 시스템 | Pretendard(KR) + Inter(숫자), `font-variant-numeric: tabular-nums` |

---

## 4. P3 — 장기 과제 (Nice-to-have)

| # | 항목 | 설명 |
|---|------|------|
| 1 | PWA Service Worker + manifest.json | 하이브리드 캐싱 (API=network-first, 정적=cache-first) |
| 2 | ESG 통합 | KCGS(한국) + Finnhub/FMP(글로벌), 10-15% 리스크 오버레이 |
| 3 | 온보딩 플로우 | 초보/중급/상급 선택 → 첫 종목 추가 → 즉시 가치 전달 |
| 4 | KakaoTalk 공유 | 익명화 수익률 스크린샷, Web Share API |
| 5 | 시나리오/벤치마크 비교 | 현재 vs 제안 포트폴리오 side-by-side, What-if 시뮬레이터 |
| 6 | CSV/PDF 내보내기 | 포트폴리오 리포트, 브로커별 CSV 임포트 템플릿 |
| 7 | 팩터 가중치 캘리브레이션 | IC 모멘텀 기반 반기별 동적 조정 |

---

## 5. v5.5에서 완료된 수정 (참고용)

> 아래 항목들은 이미 수정 완료되었으므로 재작업 불필요.

| # | 항목 | 상태 |
|---|------|------|
| 1 | CSS `--space-*` 중복 선언 | ✅ 해결 |
| 2 | `parseInt` → `parseFloat` (소수점 단가) | ✅ 해결 |
| 3 | KR_ETF 중복 데이터 (TIGER Fn메타버스) | ✅ 1건 제거 |
| 4 | 480px 열 강제 숨김 (`nth-child(n+10)`) | ✅ 제거 (스크롤로 대체) |
| 5 | input 전역 스타일 범위 제한 | ✅ `:not()` 셀렉터 |
| 6 | `--fs-*` vs `--font-*` 이중 시스템 주석 | ✅ 의도적 분리 문서화 |
| 7 | `--color-gain/loss` 색상 규약 주석 | ✅ .positive/.negative 역할 구분 문서화 |
| 8 | 평균단가 input `step="any"` | ✅ 소수점 HTML 속성 |
| 9 | 내 포지션 hero/입력폼/모달 하드코딩 폰트 → CSS 변수 (8건) | ✅ 전건 |
| 10 | `editMyHolding` 임시 백업 (`_editBackup`) | ✅ 삭제 전 백업 |
| 11 | **투자 지식 탭 터치 불가 (Critical)** | ✅ `initKnowledgeEvents()` 신설 |
| 12 | 버전 표기 v5.4→v5.5 | ✅ |

---

## 6. 향후 AI를 위한 작업 프롬프트 모음

### Phase 1: P0 Critical (1-2일)

```
=== TASK 1: safeDivide 유틸리티 ===
파일: index.html
목표: NaN/Infinity 전파 방지

1. 아래 함수를 <script> 시작부에 추가:
   function safeDivide(num, den, fallback) {
     if (arguments.length < 3) fallback = null;
     if (!Number.isFinite(num) || !Number.isFinite(den) || den === 0) return fallback;
     return num / den;
   }

2. 파일 내 모든 나눗셈 연산을 검색 (/, /=)하고,
   재무 비율 계산(PER, PBR, ROE, 배당수익률, 부채비율, PEG 등)에
   safeDivide()를 적용하세요.

3. 복합 스코어(TotalScore 등) 계산에서 null/NaN 메트릭이 있으면
   해당 메트릭을 가중치에서 제외하고 잔여 가중치를 재정규화하세요.

4. 테스트: screeningData에서 PER=null, ROE=null인 종목(ETF 등)의
   TotalScore가 NaN이 아닌 유효한 숫자인지 확인하세요.
```

```
=== TASK 2: 비동기 초기화 순서 보장 ===
파일: index.html, AppStorage IIFE 및 initialize() 함수
목표: IndexedDB 레이스 컨디션 제거

1. AppStorage.init()가 반환하는 Promise가 resolve된 후에만
   screeningData 접근 및 UI 렌더링이 시작되도록 보장하세요.
2. db.onversionchange = function() { db.close(); location.reload(); }; 추가
3. IDB 트랜잭션 내에서 await fetch() 등 비-IDB 비동기 호출이 없는지 확인하세요.
```

### Phase 2: P1 Major (1-2주)

```
=== TASK 3: 매도 신호 재보정 ===
파일: index.html, evaluateMarket() 함수
목표: 고정 임계값 → 레짐 인식 신호

1. VIX 신호:
   - 기존: VIX > 35 → 방어
   - 개선: VIX 252일 백분위 > 90th AND 5일 변화율 > +40% → 방어
   - 역발상: VIX 252일 백분위 > 95th → "공포 극대화, 역발상 매수 고려" 표시

2. CAPE 신호:
   - 기존: CAPE > 35 → 매도
   - 개선: CAPE를 10년 기대수익률 감쇠 함수로 사용
     expectedReturn = 1/CAPE (Shiller 역수)
     CAPE > 30 → "장기 기대수익률 3.3% 이하, 주식 비중 점진적 축소 권장"

3. HY spread 신호:
   - 기존: > 500bps → 경고
   - 추가: 30일 변화율 > +100bps → 긴급도 에스컬레이션
```

```
=== TASK 4: 세금 손실 수확 알림 ===
파일: index.html, 내 포지션 탭 renderMyHoldings() 근처
목표: 한국 해외주식 양도세 최적화 알림

1. myHoldings에서 _sheet='US_Stocks'인 포지션 필터링
2. 각 포지션의 미실현 손익 계산 (curPrice 존재 시)
3. 연간 실현 이익 합계가 ₩2,500,000 초과이고,
   미실현 손실 >₩500,000인 포지션이 있으면 AlertCard 표시:
   "📊 [종목명] 매도 시 약 ₩XX 세금 절감 가능"
4. 한국 세법: 양도세 22%, 연간 공제 ₩2,500,000, wash-sale 미적용
```

```
=== TASK 5: 가상 스크롤 ===
파일: index.html, renderTable() 함수
목표: 종목 테이블 DOM 90% 절감

1. renderTable()에서 전체 행 렌더링 → 보이는 행(+버퍼 20행)만 렌더링
2. 상단/하단에 spacer div로 스크롤바 높이 유지
3. IntersectionObserver 또는 scroll 이벤트로 보이는 범위 추적
4. requestAnimationFrame으로 DOM 업데이트 배칭
5. 기존 정렬/필터 기능이 정상 동작하는지 확인
```

### Phase 3: P2 Enhancement (2-4주)

```
=== TASK 6: Black-Litterman 배분 엔진 ===
파일: index.html (또는 별도 bl-engine.js)
목표: 기존 휴리스틱 배분을 학술 기반 모델로 교체

1. Ledoit-Wolf 축소 추정기로 공분산 행렬 추정
2. 시가총액 가중치에서 균형 수익률 역산: Π = δΣw_mkt (δ≈3.0)
3. 기존 팩터 스코어/evaluateMarket() 결과를 "뷰"로 변환:
   - TotalScore > 65 종목 → +2% 초과수익 뷰 (신뢰도 60%)
   - evaluateMarket() "방어" → 주식 -5% 뷰 (신뢰도 70%)
4. BL 공식으로 사후 기대수익률 계산 → MVO로 최적 가중치 산출
5. UI: 기존 배분 추천 옆에 "BL 최적 배분" 카드 추가,
   각 자산의 균형 수익률 vs BL 사후 수익률 비교 표시
```

```
=== TASK 7: China Hard Landing 시나리오 추가 ===
파일: index.html, knowledge-template 내 scenario 섹션
목표: 한국 경제 특수 리스크 반영

1. 시나리오 카테고리 "이벤트"에 "🇨🇳 차이나 경착륙" 시나리오 추가
2. 포트폴리오 영향 캘리브레이션:
   - 한국 주식: -20~-30% (반도체 수출 25% 의존)
   - 한국 채권: +3~5% (안전자산 선호)
   - 금: +10~15%
   - 미국 주식: -10~-15%
3. 대응 전략: SK하이닉스/삼성전자 비중 축소, 내수주/방어주 전환
4. 조건 체크리스트: 중국 PMI, 부동산 가격, 위안화 환율, 한국 반도체 수출액
5. initKnowledgeEvents() 내에서 새 시나리오 탭 이벤트 바인딩 확인
```

```
=== TASK 8: 내 포지션 탭 UX 개선 ===
파일: index.html, #myposition 섹션 + renderMyHoldings()

1. 빈 상태 UI: 보유종목 0건 시
   <div style="text-align:center;padding:60px;">
     📊 아직 등록된 종목이 없습니다
     <br><button>+ 종목 추가하기</button>
   </div>

2. 탭 전환 스크롤 복원:
   - 각 탭의 scrollTop을 변수에 저장
   - 탭 전환 시 저장된 위치로 복원

3. "모닝 브리프" 카드 (보유종목 있을 때):
   - 전일 대비 포트폴리오 변동 요약
   - 우선순위 액션 1-2개 (리밸런싱 필요, 손실 수확 기회 등)

4. 리밸런싱 알림:
   - 목표 비중 대비 200bps 이상 이탈 시 경고 카드
   - "리밸런싱 미리보기" 버튼 → 필요 매매 수량/금액 표시
```

---

## 7. 기존 데이터 이슈 (코드 수정 불가, 외부 데이터 파이프라인에서 해결)

> 이 항목들은 코드가 아닌 데이터 소스 쪽에서 수정이 필요합니다.

| # | 항목 | 상세 |
|---|------|------|
| 1 | KR_ETF Code 401470 중복 | "KODEX 미국러셀2000(H)" vs "KODEX 수소경제" 동일 코드. 후자 실제 코드: 385510 추정 |
| 2 | US_Stocks DivYield 이상값 | COF 128, PGR 646, VZ 682 등. bp 단위 혼재 추정 |
| 3 | From52wHigh 양수 7건 | KODEX 반도체MV +6.28% 등. 갱신 시점 불일치 |
| 4 | KR_Stocks에 ETF 혼재 | KODEX 200(Rank 39) 등이 주식 시트에 포함 |

---

## 8. 코드 수정 시 주의사항

1. **버전 관리**: 사소한 수정이라도 `<title>` 태그의 버전 번호를 변경할 것 (minor: 5.5→5.6, major: 5.5→6.0)
2. **CSS 변수 이중 시스템**: `--fs-*`(콘텐츠 고정)와 `--font-*`(UI 반응형)는 의도적 분리. 통합 시 반응형 깨짐
3. **한국 색상 규약**: `--color-gain`(빨강)=상승, `--color-loss`(파랑)=하락. `.positive/.negative`는 재무 건전성용 (초록/빨강)
4. **`<template id="knowledge-template">`**: 내부 콘텐츠는 DOM에 없음. 새 인터랙티브 요소 추가 시 반드시 `initKnowledgeEvents()` 내에서 이벤트 바인딩
5. **screeningData**: 단일 행 JSON. 수정 시 JSON 문법 오류 주의 (쉼표, 따옴표)
6. **AppStorage**: IndexedDB 우선, localStorage 폴백. `get()`은 동기(인메모리 캐시), `set()`은 비동기(IDB 쓰기)
7. **HTML 태그 밸런스**: div 1273쌍, details 21쌍. 수정 후 반드시 검증

---

## 부록: 미반영 사유 분류

| 분류 | 건수 | 대표 항목 |
|------|------|----------|
| 외부 데이터 필요 | 16건 | RMW, CMA, QMJ, SUE/ERR, Credit Impulse, SLOOS, ACM, VIX Term/SKEW/P-C, GPR, Amihud, ESG, FCF, Markov, FX Forward, 배당히스토리, 거래대금 |
| 대규모 아키텍처 변경 | 7건 | Black-Litterman, True ERC, TSMOM, Virtual Scrolling, Service Worker, 시나리오 확률, 팩터가중치 재구조 |
| 별도 UI 구현 | 4건 | 온보딩, KakaoTalk, 벤치마크 비교, China 시나리오 |
| 한국 세법 로직 | 1건 | Tax-loss harvesting |

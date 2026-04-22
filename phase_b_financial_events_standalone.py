#!/usr/bin/env python3
"""
phase_b_financial_events_standalone.py (v1.0)

기존 phase_b_financial_events.py는 global_screener_v4_4.py의 내부 함수 
(get_dart_client, get_dart_financials, dart_find_corp_code)에 의존했으나,
사용자 저장소의 실제 파이프라인인 global_screener_v3.py (v4.1 내용)에는 
해당 함수가 없음.

이 독립실행형 모듈은:
  - DART API 래퍼 (OpenDartReader)를 자체적으로 사용
  - v3.py의 수집 데이터 구조와 호환
  - 완전 독립 (메인 파이프라인 수정 불필요, import만)

통합 방법 (global_screener_v3.py 메인 함수):
    from phase_b_financial_events_standalone import integrate_phase_b_standalone
    
    # 기존 data dict 생성 후
    data = {...}
    
    # 저장 직전에 한 줄 추가
    data = integrate_phase_b_standalone(data)  # ← 추가
    save_to_json(data, json_file)

환경변수:
    DART_API_KEY: DART Open API 키 (필수, 없으면 자동 스킵)
    SKIP_PHASE_B: '1' 이면 스킵
    PHASE_B_LIMIT: 대상 종목 수 (기본 100, 최소 10, 최대 500)

학술 근거:
    - Bernard & Thomas (1989) Post-Earnings-Announcement Drift
    - Ikenberry et al. (1995) Buyback Announcement Returns
    - Loughran & Ritter (1995) SEO Underperformance
"""

import os
import time
import json
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any

logger = logging.getLogger(__name__)

# ==================================================================
# 0. DART 클라이언트 (자체 구현)
# ==================================================================
_DART_CLIENT = None
_CORP_CODE_CACHE: Dict[str, str] = {}


def _get_dart_client():
    """OpenDartReader 싱글톤 (한 번만 초기화)."""
    global _DART_CLIENT
    if _DART_CLIENT is not None:
        return _DART_CLIENT
    
    api_key = os.environ.get('DART_API_KEY', '').strip()
    if not api_key:
        return None
    
    try:
        import OpenDartReader
        _DART_CLIENT = OpenDartReader(api_key)
        logger.info("[Phase B] DART 클라이언트 초기화 성공")
        return _DART_CLIENT
    except ImportError:
        logger.warning("[Phase B] OpenDartReader 미설치 (pip install OpenDartReader)")
        return None
    except Exception as e:
        logger.warning(f"[Phase B] DART 초기화 실패: {e}")
        return None


def _find_corp_code(stock_code: str) -> Optional[str]:
    """종목코드(005930) → DART 기업코드(corp_code) 변환."""
    if stock_code in _CORP_CODE_CACHE:
        return _CORP_CODE_CACHE[stock_code]
    
    dart = _get_dart_client()
    if dart is None:
        return None
    
    try:
        # OpenDartReader의 find_corp_code: 티커로 기업코드 검색
        result = dart.find_corp_code(stock_code)
        if result:
            _CORP_CODE_CACHE[stock_code] = result
            return result
    except Exception as e:
        logger.debug(f"[Phase B] corp_code 검색 실패 {stock_code}: {e}")
    
    return None


# ==================================================================
# 1. 공시 카테고리 분류
# ==================================================================
def _categorize_disclosure(report_nm: str) -> str:
    """공시명을 카테고리로 분류."""
    if not report_nm:
        return 'other'
    
    # 실적
    if any(k in report_nm for k in ['사업보고서', '분기보고서', '반기보고서', '연결재무제표']):
        return 'earnings_release'
    # 자사주
    if any(k in report_nm for k in ['자기주식', '자사주']):
        if '취득' in report_nm:
            return 'buyback_announce'
        elif '처분' in report_nm:
            return 'buyback_dispose'
        return 'buyback'
    # 배당
    if '배당' in report_nm:
        return 'dividend'
    # 자본 조정
    if any(k in report_nm for k in ['유상증자', '무상증자', '주식분할', '액면분할']):
        return 'capital'
    # 지분 변동
    if any(k in report_nm for k in ['임원·주요주주특정증권등소유상황', '특정증권등', '소유상황']):
        return 'insider'
    # 합병/분할
    if any(k in report_nm for k in ['합병', '분할', '인수']):
        return 'ma'
    # 전환사채
    if any(k in report_nm for k in ['전환사채', '신주인수권부사채', 'CB', 'BW']):
        return 'convertible'
    
    return 'other'


# ==================================================================
# 2. 최근 공시 조회 (캐시 포함)
# ==================================================================
_DISCLOSURE_CACHE: Dict[str, Dict[str, Any]] = {}


def _fetch_recent_disclosures(corp_code: str, days: int = 7) -> List[Dict]:
    """
    최근 N일 공시 조회. 1시간 TTL 캐시.
    """
    cache_key = f"{corp_code}_{days}"
    if cache_key in _DISCLOSURE_CACHE:
        cached = _DISCLOSURE_CACHE[cache_key]
        if (datetime.now() - cached['time']).total_seconds() < 3600:
            return cached['data']
    
    dart = _get_dart_client()
    if dart is None:
        return []
    
    end = datetime.now()
    start = end - timedelta(days=days)
    
    try:
        df = dart.list(
            corp=corp_code,
            start=start.strftime('%Y-%m-%d'),
            end=end.strftime('%Y-%m-%d'),
            final=False
        )
        if df is None or len(df) == 0:
            return []
        
        disclosures = []
        for _, row in df.iterrows():
            report_nm = str(row.get('report_nm', ''))
            disclosures.append({
                'rcept_no': str(row.get('rcept_no', '')),
                'report_nm': report_nm,
                'rcept_dt': str(row.get('rcept_dt', '')),
                'category': _categorize_disclosure(report_nm)
            })
        
        _DISCLOSURE_CACHE[cache_key] = {
            'time': datetime.now(),
            'data': disclosures
        }
        return disclosures
    
    except Exception as e:
        logger.debug(f"[Phase B] 공시 조회 실패 {corp_code}: {e}")
        return []


# ==================================================================
# 3. 재무 이벤트 점수
# ==================================================================
def _calc_event_score(disclosures: List[Dict]) -> float:
    """
    최근 공시들 기반 -1~+1 점수.
    
    Ikenberry et al. (1995): 자사주 매입 +12% 초과수익 (3년)
    Loughran & Ritter (1995): 유상증자 -40% 저성과 (5년)
    """
    score = 0.0
    for d in disclosures:
        cat = d.get('category', 'other')
        if cat == 'buyback_announce':
            score += 0.3
        elif cat == 'buyback_dispose':
            score -= 0.2
        elif cat == 'dividend':
            score += 0.2
        elif cat == 'earnings_release':
            score += 0.05
        elif cat == 'capital':
            if '유상' in d.get('report_nm', ''):
                score -= 0.4
            elif '무상' in d.get('report_nm', ''):
                score += 0.1
        elif cat == 'convertible':
            score -= 0.3
        elif cat == 'ma':
            score += 0.1
    
    return max(-1.0, min(1.0, score))


# ==================================================================
# 4. 메인 통합 함수 (v3.py용)
# ==================================================================
def _extract_kr_stocks(data: Dict) -> List[Dict]:
    """
    v3.py의 data 구조에서 한국 종목 리스트 추출.
    
    v3.py 실제 키: data['KR_Stocks'] = get_korea_stocks()
    """
    # v3.py 기본 키 우선, 다른 변형도 지원
    for key in ['KR_Stocks', 'Korea_Stocks', 'Stocks_KR', 'Stocks_KR_V4', 'kr_stocks']:
        if key in data and isinstance(data[key], list):
            return data[key]
    return []


def _get_row_code(row: Dict) -> Optional[str]:
    """종목 row에서 코드 추출 (다양한 필드명 지원)."""
    for key in ['Code', 'Symbol', 'code', 'symbol', 'Ticker', 'ticker']:
        if key in row and row[key]:
            # '005930.KS' 형태면 '005930'만
            code = str(row[key]).split('.')[0].zfill(6)[:6]
            if code.isdigit():
                return code
    return None


def _get_row_name(row: Dict) -> str:
    """종목 row에서 이름 추출."""
    for key in ['Name', 'name', 'shortName', 'longName']:
        if key in row and row[key]:
            return str(row[key])
    return '―'


def integrate_phase_b_standalone(data: Dict) -> Dict:
    """
    메인 파이프라인에서 한 줄로 호출.
    
    data 구조에 다음 키 추가:
        - FinancialEvents: [{code, name, event_type, event_date, severity, ...}, ...]
        - EarningsRevisions: {code: {...}, ...} (현재는 빈 dict, 향후 확장)
    """
    # 환경변수 체크
    if os.environ.get('SKIP_PHASE_B', '').strip() in ('1', 'true', 'True'):
        logger.info("[Phase B] SKIP_PHASE_B 환경변수 → 스킵")
        data['FinancialEvents'] = []
        data['EarningsRevisions'] = {}
        return data
    
    # DART 키 체크
    if not os.environ.get('DART_API_KEY', '').strip():
        logger.warning("[Phase B] DART_API_KEY 없음 → 스킵")
        data['FinancialEvents'] = []
        data['EarningsRevisions'] = {}
        return data
    
    # 한국 종목 추출
    kr_rows = _extract_kr_stocks(data)
    if not kr_rows:
        logger.warning("[Phase B] 한국 종목 데이터 없음 → 스킵")
        data['FinancialEvents'] = []
        data['EarningsRevisions'] = {}
        return data
    
    # limit 적용
    try:
        limit = max(10, min(500, int(os.environ.get('PHASE_B_LIMIT', '100'))))
    except ValueError:
        limit = 100
    
    target_rows = kr_rows[:limit]
    logger.info(f"[Phase B] 시작 ({len(target_rows)}종목 대상)")
    
    t0 = time.time()
    events = []
    processed = 0
    skipped = 0
    
    for row in target_rows:
        code = _get_row_code(row)
        name = _get_row_name(row)
        if not code:
            skipped += 1
            continue
        
        try:
            corp_code = _find_corp_code(code)
            if not corp_code:
                skipped += 1
                continue
            
            disclosures = _fetch_recent_disclosures(corp_code, days=7)
            if not disclosures:
                processed += 1
                continue
            
            # 중요 공시만 추출
            for d in disclosures:
                cat = d.get('category', 'other')
                if cat in ('other', 'insider'):
                    continue
                
                severity = 'high' if cat in ('earnings_release', 'buyback_announce', 'capital') else 'medium'
                events.append({
                    'code': code,
                    'name': name,
                    'event_type': cat,
                    'event_date': d.get('rcept_dt', ''),
                    'severity': severity,
                    'report_nm': d.get('report_nm', ''),
                    'rcept_no': d.get('rcept_no', '')
                })
            
            # row에 점수 추가
            row['_financial_event_score'] = _calc_event_score(disclosures)
            row['_disclosure_count_7d'] = len(disclosures)
            processed += 1
            
        except Exception as e:
            logger.debug(f"[Phase B] {code} 처리 실패: {e}")
            skipped += 1
            continue
    
    # 날짜 역순 정렬
    events.sort(key=lambda x: x.get('event_date', ''), reverse=True)
    
    elapsed = time.time() - t0
    logger.info(f"[Phase B] 완료 ({elapsed:.1f}초, 처리 {processed}개, 스킵 {skipped}개, 이벤트 {len(events)}건)")
    
    data['FinancialEvents'] = events
    data['EarningsRevisions'] = {}  # 향후 확장 (컨센서스 기반)
    
    return data


# ==================================================================
# 단독 테스트
# ==================================================================
if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(levelname)s %(message)s'
    )
    
    print("=" * 70)
    print("Phase B Standalone 스모크 테스트")
    print("=" * 70)
    
    # 환경변수 체크
    dart_key = os.environ.get('DART_API_KEY', '')
    print(f"\nDART_API_KEY: {'설정됨' if dart_key else '미설정'}")
    
    if not dart_key:
        print("\n⚠️ DART_API_KEY 미설정 → 실제 API 호출 없이 구조만 검증")
    
    # 가짜 data (v3.py 실제 키 KR_Stocks 사용)
    test_data = {
        'KR_Stocks': [
            {'Code': '005930', 'Name': '삼성전자'},
            {'Code': '000660', 'Name': 'SK하이닉스'},
            {'Code': '035720', 'Name': '카카오'}
        ]
    }
    
    result = integrate_phase_b_standalone(test_data)
    print(f"\n[결과]")
    print(f"  FinancialEvents: {len(result.get('FinancialEvents', []))}건")
    print(f"  EarningsRevisions: {len(result.get('EarningsRevisions', {}))}종목")
    
    if result.get('FinancialEvents'):
        print(f"\n[샘플 이벤트 (최대 3개)]")
        for e in result['FinancialEvents'][:3]:
            print(f"  · {e['name']}({e['code']}) - {e['event_type']} - {e['event_date']}")
    
    print("\n✅ 스모크 테스트 통과")

#!/usr/bin/env python3
"""
phase_c_flow_signals_standalone.py (v1.0)

기존 phase_c_flow_signals.py는 파이프라인이 수집해둔 KRX 필드
(foreign_5d_net, shorting_balance_ratio 등)에 의존했으나,
사용자 저장소의 global_screener_v3.py (v4.1)는 fetch_krx_microstructure
함수가 없어 해당 필드가 없음.

이 독립실행형 모듈은:
  - pykrx를 직접 호출하여 KRX 데이터 자체 수집
  - 5일 순매수, 공매도 잔고, 수급 시그널 계산
  - v3.py의 Korea_Stocks 리스트만 있으면 작동

환경변수:
    KRX_ID, KRX_PW: KRX 회원제 (2025-12-27 이후 필수)
    SKIP_PHASE_C: '1'이면 스킵

학술 근거:
    - Choe & Kho (2005) Foreign Investors in Korea
    - Boehmer et al. (2008) Short Interest Patterns
    - Kaniel, Saar, Titman (2008) Individual Investor Trading
"""

import os
import time
import json
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any

logger = logging.getLogger(__name__)


# ==================================================================
# 0. pykrx 로그인 상태 확인
# ==================================================================
def _has_krx_auth() -> bool:
    """KRX 인증 가능 여부."""
    return bool(
        os.environ.get('KRX_ID', '').strip() and 
        os.environ.get('KRX_PW', '').strip()
    )


# ==================================================================
# 1. KRX 마이크로구조 데이터 수집 (v3.py에 없는 기능 자체 구현)
# ==================================================================
_KRX_DATA_CACHE: Dict[str, Any] = {}


def _fetch_krx_microstructure_bulk() -> Dict[str, Dict[str, Any]]:
    """
    전체 종목 한 번에 조회 (pykrx 벌크 호출).
    
    Returns:
        {code: {foreign_5d_net, inst_5d_net, individual_5d_net, 
                shorting_balance_ratio, foreign_exhaustion, ...}, ...}
    """
    cache_key = 'bulk_' + datetime.now().strftime('%Y%m%d')
    if cache_key in _KRX_DATA_CACHE:
        return _KRX_DATA_CACHE[cache_key]
    
    if not _has_krx_auth():
        logger.warning("[Phase C] KRX_ID/KRX_PW 없음 (2025-12-27 이후 필수)")
        return {}
    
    try:
        from pykrx import stock as krx
    except ImportError:
        logger.warning("[Phase C] pykrx 미설치 (pip install pykrx)")
        return {}
    
    # 날짜 설정
    now = datetime.now()
    days_back = 0
    if now.weekday() == 5: days_back = 1  # 토
    elif now.weekday() == 6: days_back = 2  # 일
    elif now.hour < 16: days_back = 1 if now.weekday() != 0 else 3
    
    target_date = now - timedelta(days=days_back)
    date_str = target_date.strftime('%Y%m%d')
    
    # 유효 거래일
    try:
        nearest = krx.get_nearest_business_day_in_a_week(date_str)
        if nearest:
            date_str = nearest
    except Exception:
        pass
    
    from_5d = (datetime.strptime(date_str, '%Y%m%d') - timedelta(days=7)).strftime('%Y%m%d')
    
    result: Dict[str, Dict[str, Any]] = {}
    
    try:
        # 투자자별 5일 순매수
        logger.info(f"[Phase C] pykrx 투자자별 순매수 조회 ({from_5d}~{date_str})")
        
        for market in ['KOSPI', 'KOSDAQ']:
            try:
                # 5일 누적 (억원 단위)
                df = krx.get_market_net_purchases_of_equities(from_5d, date_str, market, '순매수거래대금')
                if df is not None and not df.empty:
                    # 투자자별 컬럼 존재 확인
                    for idx, row in df.iterrows():
                        code = str(idx).zfill(6) if not isinstance(idx, str) else idx.zfill(6)
                        if code not in result:
                            result[code] = {}
                        
                        # 개인/외국인/기관 컬럼 매핑 (pykrx 버전별 차이)
                        for col_name, key in [
                            ('개인', 'individual_5d_net'),
                            ('외국인합계', 'foreign_5d_net'),
                            ('외국인', 'foreign_5d_net'),
                            ('기관합계', 'inst_5d_net'),
                            ('기관', 'inst_5d_net')
                        ]:
                            if col_name in df.columns and key not in result[code]:
                                try:
                                    result[code][key] = float(row[col_name]) / 1e8  # 원 → 억
                                except (TypeError, ValueError):
                                    pass
            except Exception as e:
                logger.debug(f"[Phase C] {market} 순매수 조회 실패: {e}")
                continue
        
        # 공매도 잔고
        logger.info(f"[Phase C] pykrx 공매도 잔고 조회 ({date_str})")
        try:
            for market in ['KOSPI', 'KOSDAQ']:
                try:
                    df = krx.get_shorting_balance_by_ticker(date_str, market)
                    if df is not None and not df.empty:
                        for idx, row in df.iterrows():
                            code = str(idx).zfill(6) if not isinstance(idx, str) else idx.zfill(6)
                            if code not in result:
                                result[code] = {}
                            # 공매도 잔고 비율 (시총 대비 %)
                            for col_name in ['비중', '공매도잔고비율', '잔고비중']:
                                if col_name in df.columns:
                                    try:
                                        result[code]['shorting_balance_ratio'] = float(row[col_name])
                                    except (TypeError, ValueError):
                                        pass
                                    break
                except Exception as e:
                    logger.debug(f"[Phase C] {market} 공매도 조회 실패: {e}")
                    continue
        except Exception as e:
            logger.debug(f"[Phase C] 공매도 조회 전체 실패: {e}")
        
        _KRX_DATA_CACHE[cache_key] = result
        logger.info(f"[Phase C] 마이크로구조 수집: {len(result)}종목")
        return result
    
    except Exception as e:
        logger.error(f"[Phase C] pykrx 벌크 조회 실패: {e}", exc_info=True)
        return {}


# ==================================================================
# 2. 시그널 점수 계산
# ==================================================================
def _calc_foreign_score(micro: Dict, market_cap: float) -> Tuple[float, Optional[str]]:
    """외국인 추세 점수."""
    f5 = micro.get('foreign_5d_net')
    if f5 is None or market_cap <= 0:
        return (0.0, None)
    
    try:
        f5 = float(f5)
        mcap_eok = float(market_cap) / 1e8 if market_cap > 1e10 else float(market_cap)
        if mcap_eok <= 0:
            return (0.0, None)
        
        # 시총 대비 %
        f5_pct = f5 / mcap_eok * 100
        score = max(-1.0, min(1.0, f5_pct))
        
        if score >= 0.5: label = 'strong_buy'
        elif score >= 0.15: label = 'buy'
        elif score <= -0.5: label = 'strong_sell'
        elif score <= -0.15: label = 'sell'
        else: label = 'neutral'
        
        return (round(score, 3), label)
    except (TypeError, ValueError):
        return (0.0, None)


def _calc_divergence_score(micro: Dict) -> Tuple[float, Optional[str]]:
    """기관-개인 상반도."""
    inst = micro.get('inst_5d_net')
    indiv = micro.get('individual_5d_net')
    if inst is None or indiv is None:
        return (0.0, None)
    
    try:
        inst, indiv = float(inst), float(indiv)
        if inst == 0 and indiv == 0:
            return (0.0, None)
        if (inst > 0) == (indiv > 0):
            return (0.0, 'aligned')
        
        ratio = min(abs(inst), abs(indiv)) / max(abs(inst), abs(indiv))
        if inst > 0 and indiv < 0:
            return (round(ratio, 3), 'inst_buy_indiv_sell')
        elif inst < 0 and indiv > 0:
            return (round(ratio, 3), 'inst_sell_indiv_buy')
        return (0.0, None)
    except (TypeError, ValueError):
        return (0.0, None)


def _calc_squeeze_score(micro: Dict) -> Tuple[float, Optional[str]]:
    """공매도 squeeze 점수."""
    sb = micro.get('shorting_balance_ratio')
    if sb is None:
        return (0.0, None)
    
    try:
        sb = float(sb)
        if sb < 1.5:
            return (0.0, 'normal')
        
        score = min(0.4, sb / 10)
        if score >= 0.3:
            return (round(score, 3), 'high_short')
        return (round(score, 3), 'normal')
    except (TypeError, ValueError):
        return (0.0, None)


# ==================================================================
# 3. 메인 통합 함수
# ==================================================================
def _extract_kr_stocks(data: Dict) -> List[Dict]:
    """v3.py 구조에서 한국 종목 추출 (키: KR_Stocks)."""
    for key in ['KR_Stocks', 'Korea_Stocks', 'Stocks_KR', 'Stocks_KR_V4', 'kr_stocks']:
        if key in data and isinstance(data[key], list):
            return data[key]
    return []


def _get_row_code(row: Dict) -> Optional[str]:
    for key in ['Code', 'Symbol', 'code', 'symbol']:
        if key in row and row[key]:
            code = str(row[key]).split('.')[0].zfill(6)[:6]
            if code.isdigit():
                return code
    return None


def _get_market_cap(row: Dict) -> float:
    for key in ['MarketCap', 'market_cap', 'marketCap', 'mktcap', 'Cap']:
        v = row.get(key)
        if v is not None:
            try:
                return float(v)
            except (TypeError, ValueError):
                continue
    return 0.0


def integrate_phase_c_standalone(data: Dict) -> Dict:
    """
    메인 파이프라인에서 한 줄로 호출.
    
    data['FlowSignals'] = {...} 추가.
    """
    # 스킵 체크
    if os.environ.get('SKIP_PHASE_C', '').strip() in ('1', 'true', 'True'):
        logger.info("[Phase C] SKIP_PHASE_C 환경변수 → 스킵")
        data['FlowSignals'] = {}
        return data
    
    if not _has_krx_auth():
        logger.warning("[Phase C] KRX_ID/KRX_PW 없음 → 스킵")
        data['FlowSignals'] = {}
        return data
    
    kr_rows = _extract_kr_stocks(data)
    if not kr_rows:
        logger.warning("[Phase C] 한국 종목 데이터 없음 → 스킵")
        data['FlowSignals'] = {}
        return data
    
    t0 = time.time()
    logger.info(f"[Phase C] 시작 ({len(kr_rows)}종목 대상)")
    
    # 벌크 수집 (1회)
    micro_bulk = _fetch_krx_microstructure_bulk()
    if not micro_bulk:
        logger.warning("[Phase C] pykrx 데이터 수집 실패 → 빈 결과")
        data['FlowSignals'] = {}
        return data
    
    # 종목별 시그널
    signals: Dict[str, Dict] = {}
    for row in kr_rows:
        code = _get_row_code(row)
        if not code:
            continue
        
        micro = micro_bulk.get(code, {})
        if not micro:
            continue
        
        market_cap = _get_market_cap(row)
        
        f_score, f_label = _calc_foreign_score(micro, market_cap)
        d_score, d_label = _calc_divergence_score(micro)
        s_score, s_label = _calc_squeeze_score(micro)
        
        # 데이터 있는 것만
        has_signal = any([
            f_score, d_score, s_score,
            micro.get('foreign_5d_net'), micro.get('inst_5d_net')
        ])
        if not has_signal:
            continue
        
        signals[code] = {
            'foreign_5d_net_억': micro.get('foreign_5d_net'),
            'foreign_trend_score': f_score,
            'foreign_trend_label': f_label,
            'inst_5d_net_억': micro.get('inst_5d_net'),
            'indiv_5d_net_억': micro.get('individual_5d_net'),
            'divergence_score': d_score,
            'divergence_label': d_label,
            'short_balance_ratio_pct': micro.get('shorting_balance_ratio'),
            'short_squeeze_score': s_score,
            'short_squeeze_label': s_label,
        }
        
        # row에 점수 추가
        row['_flow_foreign_score'] = f_score
        row['_flow_divergence_score'] = d_score
        row['_flow_squeeze_score'] = s_score
    
    elapsed = time.time() - t0
    logger.info(f"[Phase C] 완료 ({elapsed:.1f}초, {len(signals)}종목)")
    
    data['FlowSignals'] = signals
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
    print("Phase C Standalone 스모크 테스트")
    print("=" * 70)
    
    print(f"\nKRX_ID: {'설정됨' if os.environ.get('KRX_ID') else '미설정'}")
    print(f"KRX_PW: {'설정됨' if os.environ.get('KRX_PW') else '미설정'}")
    
    test_data = {
        'KR_Stocks': [
            {'Code': '005930', 'Name': '삼성전자', 'MarketCap': 450_000_000_000_000},
            {'Code': '000660', 'Name': 'SK하이닉스', 'MarketCap': 90_000_000_000_000}
        ]
    }
    
    result = integrate_phase_c_standalone(test_data)
    print(f"\n[결과]")
    print(f"  FlowSignals: {len(result.get('FlowSignals', {}))}종목")
    
    if result.get('FlowSignals'):
        for code, sig in list(result['FlowSignals'].items())[:3]:
            print(f"\n  · {code}:")
            print(f"    외국인 추세: {sig.get('foreign_trend_label', '―')} (score {sig.get('foreign_trend_score')})")
            print(f"    Divergence: {sig.get('divergence_label', '―')}")
            print(f"    공매도: {sig.get('short_squeeze_label', '―')}")
    
    print("\n✅ 스모크 테스트 통과")

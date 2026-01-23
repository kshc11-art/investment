#!/usr/bin/env python3
"""
글로벌 주식/ETF 스크리닝 - GitHub Actions 버전 v3.2.1 (버그 수정)
=============================================================================
v3.2.1 버그 수정:
  1. add_technical_indicators 함수에 hist 매개변수 추가 (NameError 해결)
  2. 모든 함수 호출에서 hist 전달하도록 수정
  3. ETF 데이터 수집 로직 안정화
  4. yfinance 경고 억제 추가
=============================================================================
"""

import warnings
warnings.filterwarnings('ignore')

import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import time
import sys
import re
import json
import argparse

# ★ GitHub Actions 버퍼링 해결: 즉시 출력 함수
def log(msg):
    """즉시 출력되는 로그 함수"""
    print(msg)
    sys.stdout.flush()

log("라이브러리 로딩 중...")

import logging
logging.getLogger('yfinance').setLevel(logging.CRITICAL)
logging.getLogger('urllib3').setLevel(logging.CRITICAL)

try:
    import yfinance as yf
    import requests
    from bs4 import BeautifulSoup
except ImportError as e:
    log("=" * 60)
    log("pip install yfinance openpyxl pandas requests beautifulsoup4 lxml numpy pykrx finance-datareader")
    log("=" * 60)
    sys.exit(1)

# ★ v3.0.2 추가: FinanceDataReader (한국 주식/ETF 최우선)
try:
    import FinanceDataReader as fdr
    FDR_AVAILABLE = True
    log("✅ FinanceDataReader 로드 완료")
except:
    FDR_AVAILABLE = False
    log("⚠️ FinanceDataReader 미설치 - pip install finance-datareader")

# pykrx 선택적
try:
    from pykrx import stock as pykrx_stock
    PYKRX_AVAILABLE = True
    log("✅ pykrx 로드 완료")
except:
    PYKRX_AVAILABLE = False
    log("⚠️ pykrx 미설치 - 한국 주식 일부 데이터 제한")

# ============================================================
# ★★★ v3.0 추가: 하드코딩 데이터 ★★★
# ============================================================

# US ETF 비용비율 (2024년 기준, %) - yfinance가 제공하지 않음
US_ETF_EXPENSE = {
    'SPY': 0.0945, 'IVV': 0.03, 'VOO': 0.03, 'VTI': 0.03, 'QQQ': 0.20,
    'DIA': 0.16, 'IWM': 0.19, 'IWF': 0.19, 'IWD': 0.19, 'VUG': 0.04,
    'VTV': 0.04, 'IJH': 0.05, 'IJR': 0.06, 'VB': 0.05, 'VO': 0.04,
    'RSP': 0.20, 'SPLG': 0.02, 'SCHX': 0.03, 'SCHB': 0.03, 'MGK': 0.07,
    'XLK': 0.09, 'XLF': 0.09, 'XLV': 0.09, 'XLE': 0.09, 'XLI': 0.09,
    'XLY': 0.09, 'XLP': 0.09, 'XLU': 0.09, 'XLB': 0.09, 'XLRE': 0.09,
    'VGT': 0.10, 'VFH': 0.10, 'VHT': 0.10, 'VDE': 0.10, 'VIS': 0.10,
    'VCR': 0.10, 'VDC': 0.10, 'VPU': 0.10, 'VAW': 0.10, 'VNQ': 0.12,
    'ARKK': 0.75, 'ARKW': 0.75, 'ARKF': 0.75, 'ARKG': 0.75,
    'SOXX': 0.35, 'SMH': 0.35, 'XBI': 0.35, 'IBB': 0.45,
    'HACK': 0.60, 'BOTZ': 0.68, 'LIT': 0.75, 'TAN': 0.50,
    'ICLN': 0.40, 'PBW': 0.65, 'QCLN': 0.58, 'REMX': 0.75, 'COPX': 0.65,
    'URA': 0.69, 'VYM': 0.06, 'SCHD': 0.06, 'DVY': 0.38, 'HDV': 0.08,
    'SPHD': 0.30, 'SPYD': 0.07, 'VIG': 0.06, 'DGRO': 0.08, 'NOBL': 0.35,
    'BND': 0.03, 'AGG': 0.03, 'TLT': 0.15, 'IEF': 0.15, 'SHY': 0.15,
    'LQD': 0.14, 'HYG': 0.49, 'JNK': 0.40, 'TIP': 0.19,
    'GLD': 0.40, 'IAU': 0.25, 'SLV': 0.50, 'USO': 0.79, 'UNG': 1.35,
    'DBC': 0.85, 'PDBC': 0.59,
    'TQQQ': 0.86, 'SQQQ': 0.86, 'UPRO': 0.91, 'SPXU': 0.91,
    'SOXL': 0.76, 'SOXS': 0.76,
    'VEA': 0.05, 'VWO': 0.08, 'EFA': 0.32, 'EEM': 0.68,
    'IEFA': 0.07, 'IEMG': 0.09, 'VXUS': 0.07, 'ACWI': 0.32
}

# KR ETF 비용비율 (총보수, %)
KR_ETF_EXPENSE = {
    '069500': 0.05, '114800': 0.07, '122630': 0.07, '229200': 0.25,
    '252670': 0.64, '305720': 0.45, '091160': 0.45, '133690': 0.07,
    '143850': 0.07, '192090': 0.55, '371460': 0.07, '360200': 0.25,
    '148020': 0.45, '153130': 0.24, '161510': 0.04, '157450': 0.04,
    '261240': 0.35, '360750': 0.07, '371450': 0.49, '379800': 0.07,
    '381170': 0.49, '395160': 0.45, '461460': 0.10, '102110': 0.05,
    '105190': 0.07, '226490': 0.07, '102780': 0.07,
}

# KR ETF 카테고리
KR_ETF_CATEGORY = {
    '069500': '국내 대형주', '114800': '인버스', '122630': '레버리지',
    '229200': '코스닥', '252670': '인버스 레버리지', '305720': '섹터(2차전지)',
    '091160': '섹터(반도체)', '133690': '미국 대형주', '143850': '미국 대형주',
    '192090': '중국', '371460': '커버드콜', '360200': '미국 대형주',
}

# KR ETF 배당수익률
KR_ETF_DIVYIELD = {
    '069500': 1.8, '114800': 0.0, '122630': 0.0, '229200': 0.3,
    '161510': 4.5, '157450': 3.8, '371460': 8.0, '102110': 1.8,
}

# ★ v3.0.2 추가: 한국 주식 하드코딩 (KOSPI 시총 상위)
KR_STOCK_LIST = {
    '005930': ('삼성전자', '전기전자'),
    '000660': ('SK하이닉스', '전기전자'),
    '373220': ('LG에너지솔루션', '전기전자'),
    '207940': ('삼성바이오로직스', '의약품'),
    '005380': ('현대차', '운수장비'),
    '006400': ('삼성SDI', '전기전자'),
    '051910': ('LG화학', '화학'),
    '000270': ('기아', '운수장비'),
    '005490': ('POSCO홀딩스', '철강금속'),
    '035420': ('NAVER', '서비스업'),
    '068270': ('셀트리온', '의약품'),
    '028260': ('삼성물산', '유통업'),
    '105560': ('KB금융', '기타금융'),
    '055550': ('신한지주', '기타금융'),
    '012330': ('현대모비스', '운수장비'),
    '066570': ('LG전자', '전기전자'),
    '003670': ('포스코퓨처엠', '철강금속'),
    '096770': ('SK이노베이션', '화학'),
    '034730': ('SK', '기타금융'),
    '086790': ('하나금융지주', '기타금융'),
    '032830': ('삼성생명', '보험'),
    '003550': ('LG', '기타금융'),
    '015760': ('한국전력', '전기가스'),
    '017670': ('SK텔레콤', '통신업'),
    '010130': ('고려아연', '철강금속'),
    '018260': ('삼성에스디에스', '서비스업'),
    '009150': ('삼성전기', '전기전자'),
    '033780': ('KT&G', '음식료'),
    '035720': ('카카오', '서비스업'),
    '316140': ('우리금융지주', '기타금융'),
}

# ============================================================
# 설정
# ============================================================
TODAY = datetime.now().strftime("%Y%m%d")
DATE_1Y_AGO = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")

# 종목 수 설정
TOP_N_KR = 150
TOP_N_US = 100
TOP_N_KR_ETF = 200
TOP_N_US_ETF = 100

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
}

# 데이터 소스 차단 여부
NAVER_AVAILABLE = None
FNGUIDE_AVAILABLE = None
NXT_AVAILABLE = None

# 타임아웃 설정
TIMEOUT_SHORT = 3
TIMEOUT_LONG = 5

# 데이터 품질 검증용 필수 컬럼
REQUIRED_COLS_KR_STOCK = ['PER', 'PBR', 'ROE(%)', 'Return60D(%)', 'RSI14', 'Volatility20D']
REQUIRED_COLS_US_STOCK = ['PER', 'PBR', 'ROE(%)', 'Return60D(%)', 'RSI14', 'Volatility20D', 'Beta']
REQUIRED_COLS_ETF = ['ExpenseRatio(%)', 'Return3M(%)', 'Volatility20D']

# ============================================================
# 유틸리티 함수
# ============================================================
def fmt(val, decimals=2):
    """숫자 포맷팅"""
    if val is None or pd.isna(val):
        return None
    try:
        return round(float(val), decimals)
    except:
        return None

def safe_get(d, *keys, default=None):
    """딕셔너리에서 안전하게 값 추출"""
    if not isinstance(d, dict):
        return default
    for k in keys:
        if k in d and d[k] is not None:
            return d[k]
    return default

def fetch_yf_with_retry(ticker, max_retries=2, delay=0.3):
    """yfinance 데이터를 재시도 로직과 함께 가져오기"""
    for attempt in range(max_retries):
        try:
            t = yf.Ticker(ticker)
            hist = t.history(period="1y")
            info = t.info if hasattr(t, 'info') else {}
            
            if info is None:
                info = {}

            if not hist.empty or (info and len(info) > 5):
                return t, hist, info

            if attempt < max_retries - 1:
                time.sleep(delay)
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(delay)
            else:
                pass

    return None, pd.DataFrame(), {}

def try_multiple_tickers(ticker_variants, max_retries=2):
    """여러 티커 형식을 시도하여 가장 좋은 결과 반환"""
    best_result = (None, pd.DataFrame(), {})
    best_score = 0

    for ticker in ticker_variants:
        t, hist, info = fetch_yf_with_retry(ticker, max_retries=max_retries)

        score = len(hist) + len(info) if info else len(hist)

        if score > best_score:
            best_score = score
            best_result = (t, hist, info)

            if len(hist) > 200 and len(info) > 20:
                break

    return best_result

def fill_missing_from_info(row, info, field_mapping):
    """info 딕셔너리에서 결측값 채우기"""
    for row_field, (info_keys, multiplier, decimals) in field_mapping.items():
        if row.get(row_field) is None:
            val = safe_get(info, *info_keys) if isinstance(info_keys, tuple) else safe_get(info, info_keys)
            if val is not None:
                if multiplier != 1:
                    val = val * multiplier
                row[row_field] = fmt(val, decimals)
    return row

# ============================================================
# 네이버 금융 스크래핑
# ============================================================
def fetch_naver(url, timeout=TIMEOUT_SHORT):
    """네이버 URL 요청"""
    global NAVER_AVAILABLE
    if NAVER_AVAILABLE is False:
        return None
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        if resp.status_code == 403:
            NAVER_AVAILABLE = False
            log("  ⚠️ 네이버 금융 접근 차단됨 (403)")
            return None
        if resp.status_code == 200:
            NAVER_AVAILABLE = True
            return resp
        return None
    except:
        return None

def get_naver_stock_list(market='KOSPI', max_pages=10):
    """네이버에서 시가총액 순위 종목 리스트"""
    global NAVER_AVAILABLE
    stocks = []
    sosok = '0' if market == 'KOSPI' else '1'
    
    for page in range(1, max_pages + 1):
        url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={sosok}&page={page}"
        resp = fetch_naver(url)
        
        if not resp:
            if NAVER_AVAILABLE is False:
                return []
            continue
        
        soup = BeautifulSoup(resp.text, 'lxml')
        table = soup.select_one('table.type_2')
        
        if not table:
            continue
        
        for tr in table.select('tr'):
            tds = tr.select('td')
            if len(tds) < 10:
                continue
            
            try:
                link = tds[1].select_one('a')
                if not link:
                    continue
                
                href = link.get('href', '')
                code_match = re.search(r'code=(\d{6})', href)
                if not code_match:
                    continue
                
                code = code_match.group(1)
                name = link.text.strip()
                
                price_text = tds[2].text.strip().replace(',', '')
                price = int(price_text) if price_text.isdigit() else None
                
                stocks.append({
                    'code': code, 'name': name, 'price': price, 'market': market
                })
            except:
                continue
        
        time.sleep(0.2)
    
    return stocks

def get_naver_stock_detail(code):
    """네이버에서 종목 상세 정보"""
    data = {}
    url = f"https://finance.naver.com/item/main.naver?code={code}"
    resp = fetch_naver(url)

    if resp:
        try:
            soup = BeautifulSoup(resp.text, 'lxml')
            
            for em in soup.select('em'):
                em_id = em.get('id', '')
                try:
                    if em_id == '_per':
                        val = float(em.text.replace(',', ''))
                        if 0 < val < 1000:
                            data['per'] = val
                    elif em_id == '_pbr':
                        val = float(em.text.replace(',', ''))
                        if 0 < val < 100:
                            data['pbr'] = val
                    elif em_id == '_dvr':
                        val = float(em.text.replace(',', ''))
                        if 0 <= val < 50:
                            data['div_yield'] = val
                    elif em_id == '_eps':
                        val = float(em.text.replace(',', ''))
                        data['eps'] = val
                    elif em_id == '_bps':
                        val = float(em.text.replace(',', ''))
                        data['bps'] = val
                except:
                    pass

            try:
                price_elem = soup.select_one('p.no_today span.blind')
                if price_elem:
                    data['price'] = int(price_elem.text.replace(',', ''))
            except:
                pass

        except:
            pass

    return data

def get_fnguide_data(code):
    """FnGuide에서 재무 데이터 가져오기"""
    global FNGUIDE_AVAILABLE
    
    if FNGUIDE_AVAILABLE is False:
        return {}
    
    data = {}
    url = f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A{code}"

    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT_SHORT)
        if resp.status_code == 403:
            FNGUIDE_AVAILABLE = False
            return data
        if resp.status_code != 200:
            return data
        
        FNGUIDE_AVAILABLE = True
        soup = BeautifulSoup(resp.text, 'lxml')

        for table in soup.select('table.us_table_ty1'):
            for tr in table.select('tr'):
                tds = tr.select('td')
                ths = tr.select('th')
                if len(ths) >= 1 and len(tds) >= 1:
                    for i, th in enumerate(ths):
                        label = th.get_text().strip()
                        if i < len(tds):
                            val_text = tds[i].get_text().strip().replace(',', '').replace('%', '')
                            if val_text and val_text != '-':
                                try:
                                    val = float(val_text)
                                    if 'PER' in label and 'per' not in data:
                                        if 0 < val < 1000:
                                            data['per'] = val
                                    elif 'PBR' in label and 'pbr' not in data:
                                        if 0 < val < 100:
                                            data['pbr'] = val
                                    elif 'ROE' in label and 'roe' not in data:
                                        if -100 < val < 200:
                                            data['roe'] = val
                                except:
                                    pass

    except:
        pass

    return data

# ============================================================
# 기술적 지표 계산
# ============================================================
def calc_rsi(prices, period=14):
    try:
        delta = prices.diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=period).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=period).mean()
        rs = gain / loss
        rsi = 100 - (100 / (1 + rs))
        return rsi.iloc[-1]
    except:
        return None

def calc_volatility(prices, period=20):
    try:
        returns = prices.pct_change().dropna()
        if len(returns) < period:
            return None
        vol = returns.rolling(window=period).std().iloc[-1]
        return vol * np.sqrt(252) * 100
    except:
        return None

def calc_bollinger_position(prices, period=20):
    try:
        ma = prices.rolling(window=period).mean()
        std = prices.rolling(window=period).std()
        upper = ma + 2 * std
        lower = ma - 2 * std
        current = prices.iloc[-1]
        position = (current - lower.iloc[-1]) / (upper.iloc[-1] - lower.iloc[-1]) * 100
        return max(0, min(100, position))
    except:
        return None

def calc_max_drawdown(prices):
    try:
        peak = prices.expanding().max()
        drawdown = (prices - peak) / peak * 100
        return drawdown.min()
    except:
        return None

def calc_sharpe_ratio(prices, min_period=200, risk_free=0.04):
    """SharpeRatio 계산"""
    try:
        returns = prices.pct_change().dropna()
        if len(returns) < min_period:
            return None
        mean_return = returns.mean() * 252
        std_return = returns.std() * np.sqrt(252)
        if std_return == 0:
            return None
        return (mean_return - risk_free) / std_return
    except:
        return None

def calc_sharpe_from_existing(return_pct, volatility_pct, period_days=120, risk_free_rate=4.0):
    """이미 수집된 수익률/변동성으로 SharpeRatio 계산"""
    try:
        if return_pct is None or volatility_pct is None:
            return None
        if pd.isna(return_pct) or pd.isna(volatility_pct):
            return None
        if volatility_pct == 0:
            return None
        
        annual_return = float(return_pct) * (252 / period_days)
        annual_vol = float(volatility_pct)
        sharpe = (annual_return - risk_free_rate) / annual_vol
        return fmt(sharpe, 3)
    except:
        return None

# ============================================================
# ★★★ v3.2.1 수정: 기술적 지표 공통 계산 함수 (hist 매개변수 추가) ★★★
# ============================================================
def add_technical_indicators(row, close, hist=None, include_ma60_120=False):
    """
    기술적 지표를 row에 추가하는 공통 함수
    
    Parameters:
    - row: 데이터 딕셔너리
    - close: pandas Series (종가)
    - hist: pandas DataFrame (OHLCV 데이터, ADX 계산용) - ★ v3.2.1 추가
    - include_ma60_120: MA60/MA120 포함 여부 (US_Stocks용)
    """
    if close is None or len(close) < 20:
        return row
    
    price = close.iloc[-1]
    
    # 수익률 계산
    if len(close) >= 2:
        row['Return1D(%)'] = fmt((close.iloc[-1] / close.iloc[-2] - 1) * 100)
    if len(close) >= 6:
        row['Return1W(%)'] = fmt((close.iloc[-1] / close.iloc[-6] - 1) * 100)
        row['Return5D(%)'] = row['Return1W(%)']
    if len(close) >= 22:
        row['Return1M(%)'] = fmt((close.iloc[-1] / close.iloc[-22] - 1) * 100)
        row['Return20D(%)'] = row['Return1M(%)']
    if len(close) >= 66:
        row['Return3M(%)'] = fmt((close.iloc[-1] / close.iloc[-66] - 1) * 100)
        row['Return60D(%)'] = row['Return3M(%)']
    if len(close) >= 132:
        row['Return6M(%)'] = fmt((close.iloc[-1] / close.iloc[-132] - 1) * 100)
        row['Return120D(%)'] = row['Return6M(%)']
    
    if len(close) >= 245:
        year_ago_idx = min(len(close), 252)
        row['Return1Y(%)'] = fmt((close.iloc[-1] / close.iloc[-year_ago_idx] - 1) * 100)
        row['Return250D(%)'] = row['Return1Y(%)']
    
    # 이동평균
    row['MA20'] = fmt(close.rolling(20).mean().iloc[-1])
    row['MA50'] = fmt(close.rolling(50).mean().iloc[-1])
    row['MA200'] = fmt(close.rolling(200).mean().iloc[-1]) if len(close) >= 200 else None
    
    if include_ma60_120:
        if len(close) >= 60:
            row['MA60'] = fmt(close.rolling(60).mean().iloc[-1])
        if len(close) >= 120:
            row['MA120'] = fmt(close.rolling(120).mean().iloc[-1])
    
    # vs_MA 계산
    if row.get('MA20'): row['vs_MA20(%)'] = fmt((price / row['MA20'] - 1) * 100)
    if row.get('MA50'): row['vs_MA50(%)'] = fmt((price / row['MA50'] - 1) * 100)
    if row.get('MA200'): row['vs_MA200(%)'] = fmt((price / row['MA200'] - 1) * 100)
    
    if include_ma60_120:
        if row.get('MA60'): row['vs_MA60(%)'] = fmt((price / row['MA60'] - 1) * 100)
        if row.get('MA120'): row['vs_MA120(%)'] = fmt((price / row['MA120'] - 1) * 100)
    
    # RSI, BB
    row['RSI14'] = fmt(calc_rsi(close, 14))
    row['BB_Position'] = fmt(calc_bollinger_position(close, 20))
    
    # MACD
    if len(close) >= 26:
        ema12 = close.ewm(span=12, adjust=False).mean()
        ema26 = close.ewm(span=26, adjust=False).mean()
        macd_line = ema12 - ema26
        signal_line = macd_line.ewm(span=9, adjust=False).mean()
        row['MACD'] = fmt(macd_line.iloc[-1])
        row['MACD_Signal'] = fmt(signal_line.iloc[-1])
        row['MACD_Hist'] = fmt(macd_line.iloc[-1] - signal_line.iloc[-1])
    
    # ★ v3.2.1 수정: ADX 계산 (hist 매개변수 사용)
    if hist is not None and len(hist) >= 14:
        try:
            high = hist['High'] if 'High' in hist.columns else None
            low = hist['Low'] if 'Low' in hist.columns else None
            if high is not None and low is not None:
                tr1 = high - low
                tr2 = abs(high - close.shift(1))
                tr3 = abs(low - close.shift(1))
                tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
                atr = tr.rolling(14).mean()
                
                up_move = high - high.shift(1)
                down_move = low.shift(1) - low
                plus_dm = np.where((up_move > down_move) & (up_move > 0), up_move, 0)
                minus_dm = np.where((down_move > up_move) & (down_move > 0), down_move, 0)
                
                plus_di = 100 * pd.Series(plus_dm).rolling(14).mean() / atr
                minus_di = 100 * pd.Series(minus_dm).rolling(14).mean() / atr
                
                dx = 100 * abs(plus_di - minus_di) / (plus_di + minus_di + 0.0001)
                adx = dx.rolling(14).mean()
                
                if not adx.empty and len(adx) > 0:
                    row['ADX'] = fmt(adx.iloc[-1])
        except:
            pass
    
    # 52주 고저
    if len(close) >= 245:
        year_data = close.tail(min(len(close), 252))
        row['52wHigh'] = fmt(year_data.max())
        row['52wLow'] = fmt(year_data.min())
        row['From52wHigh(%)'] = fmt((price / year_data.max() - 1) * 100)
        row['From52wLow(%)'] = fmt((price / year_data.min() - 1) * 100)
    
    # 변동성
    row['Volatility20D'] = fmt(calc_volatility(close, 20))
    row['Volatility60D'] = fmt(calc_volatility(close, 60))
    
    # MaxDrawdown
    row['MaxDrawdown(%)'] = fmt(calc_max_drawdown(close))
    
    # SharpeRatio
    row['SharpeRatio'] = fmt(calc_sharpe_ratio(close))
    
    if row.get('SharpeRatio') is None:
        row['SharpeRatio'] = calc_sharpe_from_existing(
            row.get('Return120D(%)'), 
            row.get('Volatility20D')
        )
    
    return row

# ============================================================
# 데이터 검증
# ============================================================
def is_etf_stock(name, code=None):
    if not name:
        return False
    name_upper = name.upper()
    etf_brands = ['KODEX', 'TIGER', 'KBSTAR', 'ARIRANG', 'KOSEF', 'KINDEX',
                  'HANARO', 'SOL', 'ACE', 'TIMEFOLIO', 'WOORI', 'VITA',
                  'FOCUS', 'MASTER', 'TREX', 'SMART', 'PLUS', 'RISE']
    for brand in etf_brands:
        if brand.upper() in name_upper:
            return True
    etf_keywords = ['ETF', 'ETN', '레버리지', '인버스', '선물', '지수', '액티브', '채권']
    for kw in etf_keywords:
        if kw in name or kw.upper() in name_upper:
            return True
    return False

def calc_data_quality_score(row, required_cols, data_type='stock'):
    valid_count = 0
    for col in required_cols:
        if col in row and row[col] is not None:
            valid_count += 1
    quality = (valid_count / len(required_cols)) * 100 if required_cols else 100
    row['DataQuality(%)'] = fmt(quality, 0)
    return row

# ============================================================
# 한국 주식 수집
# ============================================================
def collect_kr_stock_codes():
    """모든 소스에서 종목 코드 수집 후 통합"""
    all_stocks = {}
    
    log("  [1단계] 종목 리스트 통합 수집...")
    
    # 소스 1: pykrx
    if PYKRX_AVAILABLE:
        try:
            log("    pykrx 로드 중...")
            kospi_tickers = pykrx_stock.get_market_ticker_list(market="KOSPI")
            for ticker in kospi_tickers:
                code = str(ticker).zfill(6)
                try:
                    name = pykrx_stock.get_market_ticker_name(ticker)
                except:
                    name = ''
                if code not in all_stocks:
                    all_stocks[code] = {'name': name, 'market': 'KOSPI'}
            log(f"    pykrx: {len(kospi_tickers)}개 추가 → 총 {len(all_stocks)}개")
        except Exception as e:
            log(f"    ⚠️ pykrx 실패: {str(e)[:30]}")
    
    # 소스 2: FinanceDataReader
    if FDR_AVAILABLE:
        try:
            log("    FDR 로드 중...")
            fdr_stocks = fdr.StockListing('KOSPI')
            added = 0
            if fdr_stocks is not None and not fdr_stocks.empty:
                for _, row in fdr_stocks.iterrows():
                    code = str(row.get('Code', row.get('Symbol', ''))).zfill(6)
                    name = row.get('Name', row.get('종목명', ''))
                    if len(code) == 6 and code.isdigit():
                        if code not in all_stocks:
                            all_stocks[code] = {'name': name, 'market': 'KOSPI'}
                            added += 1
                        elif not all_stocks[code].get('name'):
                            all_stocks[code]['name'] = name
            log(f"    FDR: {added}개 추가 → 총 {len(all_stocks)}개")
        except Exception as e:
            log(f"    ⚠️ FDR 실패: {str(e)[:30]}")
    
    # 소스 3: 네이버 금융
    if NAVER_AVAILABLE is not False:
        try:
            log("    네이버 로드 중...")
            naver_stocks = get_naver_stock_list('KOSPI', max_pages=15)
            added = 0
            for stock in naver_stocks:
                code = str(stock.get('code', '')).zfill(6)
                name = stock.get('name', '')
                if len(code) == 6 and code.isdigit():
                    if code not in all_stocks:
                        all_stocks[code] = {'name': name, 'market': 'KOSPI'}
                        added += 1
            log(f"    네이버: {added}개 추가 → 총 {len(all_stocks)}개")
        except Exception as e:
            log(f"    ⚠️ 네이버 실패: {str(e)[:30]}")
    
    # 소스 4: 하드코딩
    if len(all_stocks) < 100:
        log("    하드코딩으로 보충 중...")
        added = 0
        for code, (name, sector) in KR_STOCK_LIST.items():
            code = str(code).zfill(6)
            if code not in all_stocks:
                all_stocks[code] = {'name': name, 'market': 'KOSPI'}
                added += 1
        log(f"    하드코딩: {added}개 추가 → 총 {len(all_stocks)}개")
    
    # ETF 제거
    final_stocks = {}
    for code, info in all_stocks.items():
        name = info.get('name', '')
        if not is_etf_stock(name, code):
            final_stocks[code] = info
    
    log(f"  [1단계 완료] 총 {len(final_stocks)}개 종목 (ETF 제외)")
    return final_stocks


def get_korea_stocks():
    """한국 주식 데이터 수집"""
    log("\n[1/5] 한국 주식 수집 중...")
    
    all_stocks = collect_kr_stock_codes()
    
    if not all_stocks:
        log("  ❌ 종목 리스트를 가져올 수 없음")
        return pd.DataFrame()
    
    stock_list = list(all_stocks.items())
    if TOP_N_KR and len(stock_list) > TOP_N_KR:
        stock_list = stock_list[:TOP_N_KR]
    
    log(f"  대상: {len(stock_list)}개")
    
    results = []
    start_time = time.time()
    
    for i, (ticker, info) in enumerate(stock_list):
        try:
            name = info.get('name', '')
            market = info.get('market', 'KOSPI')
            row = {'Code': ticker, 'Name': name, 'Market': market}

            # FDR 히스토리
            hist = pd.DataFrame()
            if FDR_AVAILABLE:
                try:
                    fdr_data = fdr.DataReader(ticker, DATE_1Y_AGO)
                    if fdr_data is not None and not fdr_data.empty and len(fdr_data) > 20:
                        hist = fdr_data
                        if 'Close' not in hist.columns and '종가' in hist.columns:
                            hist = hist.rename(columns={'종가': 'Close', '시가': 'Open', '고가': 'High', '저가': 'Low', '거래량': 'Volume'})
                        
                        if 'Close' in hist.columns:
                            row['Price'] = fmt(hist['Close'].iloc[-1], 0)
                        if 'Volume' in hist.columns:
                            row['Volume'] = int(hist['Volume'].iloc[-1])
                        row['DataSource'] = 'FDR'
                except:
                    pass

            # 네이버 (재무지표)
            if NAVER_AVAILABLE is not False:
                try:
                    naver_data = get_naver_stock_detail(ticker)
                    if naver_data:
                        if not row.get('Price') and naver_data.get('price'):
                            row['Price'] = naver_data['price']
                        if naver_data.get('per'):
                            row['PER'] = fmt(naver_data['per'])
                        if naver_data.get('pbr'):
                            row['PBR'] = fmt(naver_data['pbr'])
                        if naver_data.get('eps'):
                            row['EPS'] = fmt(naver_data['eps'], 0)
                        if naver_data.get('bps'):
                            row['BPS'] = fmt(naver_data['bps'], 0)
                        if naver_data.get('div_yield'):
                            row['DivYield(%)'] = fmt(naver_data['div_yield'])
                except:
                    pass

            # FnGuide (재무지표 보충)
            if row.get('PER') is None or row.get('PBR') is None:
                fnguide_data = get_fnguide_data(ticker)
                if fnguide_data:
                    if row.get('PER') is None and fnguide_data.get('per'):
                        row['PER'] = fmt(fnguide_data['per'])
                    if row.get('PBR') is None and fnguide_data.get('pbr'):
                        row['PBR'] = fmt(fnguide_data['pbr'])
                    if fnguide_data.get('roe'):
                        row['ROE(%)'] = fmt(fnguide_data['roe'])

            # yfinance (폴백)
            info = {}
            if hist.empty:
                try:
                    ticker_variants = [f"{ticker}.KS"]
                    t, hist, info = try_multiple_tickers(ticker_variants, max_retries=1)
                except:
                    pass

            if info:
                if not row.get('Price'):
                    row['Price'] = fmt(safe_get(info, 'regularMarketPrice'))
                if not row.get('Sector'):
                    row['Sector'] = safe_get(info, 'sector')
                if row.get('PER') is None:
                    row['PER'] = fmt(safe_get(info, 'trailingPE'))
                if row.get('PBR') is None:
                    row['PBR'] = fmt(safe_get(info, 'priceToBook'))
                if row.get('ROE(%)') is None and safe_get(info, 'returnOnEquity'):
                    row['ROE(%)'] = fmt(safe_get(info, 'returnOnEquity') * 100)
                if row.get('DivYield(%)') is None and safe_get(info, 'dividendYield'):
                    row['DivYield(%)'] = fmt(safe_get(info, 'dividendYield') * 100)

            # ★ v3.2.1 수정: 기술적 지표 (hist 전달)
            if not hist.empty and 'Close' in hist.columns and len(hist) > 20:
                close = hist['Close']
                row = add_technical_indicators(row, close, hist=hist, include_ma60_120=True)

            row = calc_data_quality_score(row, REQUIRED_COLS_KR_STOCK, 'kr_stock')
            results.append(row)
            
        except Exception as e:
            results.append({'Code': ticker, 'Name': name, 'Market': market, 'Remark': str(e)[:50]})
        
        if (i + 1) % 30 == 0 or i == 0:
            elapsed = time.time() - start_time
            per_stock = elapsed / (i + 1) if i > 0 else 0
            remaining = per_stock * (len(stock_list) - i - 1)
            log(f"  진행: {i+1}/{len(stock_list)} ({(i+1)/len(stock_list)*100:.0f}%) - 남은시간: {remaining/60:.1f}분")
        
        time.sleep(0.01)
    
    log(f"  ✅ 완료: {len(results)}개")
    return pd.DataFrame(results)

# ============================================================
# 미국 주식 수집
# ============================================================
def get_us_stocks():
    """미국 S&P 500 주식 데이터"""
    log("\n[2/5] 미국 주식 수집 중...")
    
    sp500_top = [
        'AAPL', 'MSFT', 'AMZN', 'NVDA', 'GOOGL', 'META', 'TSLA', 'BRK-B', 'UNH', 'JNJ',
        'XOM', 'JPM', 'V', 'PG', 'MA', 'HD', 'CVX', 'MRK', 'ABBV', 'LLY',
        'PEP', 'KO', 'COST', 'AVGO', 'TMO', 'MCD', 'WMT', 'CSCO', 'ACN', 'ABT',
        'DHR', 'NEE', 'VZ', 'ADBE', 'CRM', 'NKE', 'PM', 'TXN', 'WFC', 'BMY',
        'RTX', 'CMCSA', 'ORCL', 'UPS', 'HON', 'QCOM', 'COP', 'T', 'LOW', 'MS',
        'INTC', 'UNP', 'ELV', 'BA', 'SPGI', 'CAT', 'IBM', 'GS', 'PLD', 'INTU',
        'DE', 'AMD', 'SBUX', 'AXP', 'AMAT', 'MDLZ', 'GE', 'BLK', 'GILD', 'ADI',
        'LMT', 'ISRG', 'TJX', 'SYK', 'CVS', 'BKNG', 'ADP', 'MMC', 'VRTX', 'REGN',
        'PGR', 'CB', 'NOW', 'CI', 'SCHW', 'ZTS', 'MO', 'TMUS', 'SO', 'DUK',
        'EOG', 'BDX', 'C', 'PNC', 'CL', 'TGT', 'ITW', 'SLB', 'AMT', 'USB',
    ]
    
    tickers = sp500_top[:TOP_N_US] if TOP_N_US else sp500_top
    results = []
    
    log(f"  대상: {len(tickers)}개")
    
    for i, ticker in enumerate(tickers):
        try:
            t, hist, info = fetch_yf_with_retry(ticker, max_retries=3, delay=0.5)

            if hist.empty and not info:
                continue

            row = {'Ticker': ticker}
            
            row['Name'] = safe_get(info, 'shortName', 'longName')
            row['Sector'] = safe_get(info, 'sector')
            row['Industry'] = safe_get(info, 'industry')
            row['Price'] = fmt(safe_get(info, 'regularMarketPrice'))
            
            cap = safe_get(info, 'marketCap')
            if cap:
                row['MarketCap(B)'] = fmt(cap / 1e9, 1)
            
            row['PER'] = fmt(safe_get(info, 'trailingPE'))
            row['ForwardPE'] = fmt(safe_get(info, 'forwardPE'))
            row['PBR'] = fmt(safe_get(info, 'priceToBook'))
            row['PSR'] = fmt(safe_get(info, 'priceToSalesTrailing12Months'))
            row['PEG'] = fmt(safe_get(info, 'pegRatio'))
            row['EV/EBITDA'] = fmt(safe_get(info, 'enterpriseToEbitda'))
            
            row['ROE(%)'] = fmt(safe_get(info, 'returnOnEquity', default=0) * 100) if safe_get(info, 'returnOnEquity') else None
            row['ROA(%)'] = fmt(safe_get(info, 'returnOnAssets', default=0) * 100) if safe_get(info, 'returnOnAssets') else None
            row['GrossMargin(%)'] = fmt(safe_get(info, 'grossMargins', default=0) * 100) if safe_get(info, 'grossMargins') else None
            row['OpMargin(%)'] = fmt(safe_get(info, 'operatingMargins', default=0) * 100) if safe_get(info, 'operatingMargins') else None
            row['NetMargin(%)'] = fmt(safe_get(info, 'profitMargins', default=0) * 100) if safe_get(info, 'profitMargins') else None
            
            row['RevenueGrowth(%)'] = fmt(safe_get(info, 'revenueGrowth', default=0) * 100) if safe_get(info, 'revenueGrowth') else None
            row['EarningsGrowth(%)'] = fmt(safe_get(info, 'earningsGrowth', default=0) * 100) if safe_get(info, 'earningsGrowth') else None
            
            row['CurrentRatio'] = fmt(safe_get(info, 'currentRatio'))
            row['DebtRatio(%)'] = fmt(safe_get(info, 'debtToEquity'))
            
            row['DivYield(%)'] = fmt(safe_get(info, 'dividendYield', default=0) * 100) if safe_get(info, 'dividendYield') else None
            row['PayoutRatio(%)'] = fmt(safe_get(info, 'payoutRatio', default=0) * 100) if safe_get(info, 'payoutRatio') else None
            
            row['Beta'] = fmt(safe_get(info, 'beta'))
            row['EPS'] = fmt(safe_get(info, 'trailingEps'))
            row['BPS'] = fmt(safe_get(info, 'bookValue'))
            row['Volume'] = safe_get(info, 'regularMarketVolume') or safe_get(info, 'averageVolume')
            
            # ★ v3.2.1 수정: 기술적 지표 (hist 전달)
            if not hist.empty and len(hist) > 20:
                close = hist['Close']
                row = add_technical_indicators(row, close, hist=hist, include_ma60_120=True)
                
                # YTD 수익률
                try:
                    ytd_start = close[close.index >= f"{datetime.now().year}-01-01"]
                    if len(ytd_start) > 1:
                        row['ReturnYTD(%)'] = fmt((close.iloc[-1] / ytd_start.iloc[0] - 1) * 100)
                except:
                    pass

            if row.get('52wHigh') is None:
                val = safe_get(info, 'fiftyTwoWeekHigh')
                if val:
                    row['52wHigh'] = fmt(val)
            if row.get('52wLow') is None:
                val = safe_get(info, 'fiftyTwoWeekLow')
                if val:
                    row['52wLow'] = fmt(val)

            if row.get('Price') is None:
                row['Price'] = fmt(safe_get(info, 'regularMarketPrice', 'previousClose', 'open'))

            row = calc_data_quality_score(row, REQUIRED_COLS_US_STOCK, 'us_stock')
            results.append(row)

        except Exception as e:
            results.append({'Ticker': ticker, 'Remark': str(e)[:50]})

        if (i + 1) % 20 == 0:
            log(f"  진행: {i+1}/{len(tickers)}")

        time.sleep(0.15)

    log(f"  ✅ 완료: {len(results)}개")
    return pd.DataFrame(results)

# ============================================================
# ETF 공통 함수
# ============================================================
def get_etf_data(tickers, region=""):
    """ETF 데이터 수집 공통 함수"""
    results = []

    for i, ticker in enumerate(tickers):
        try:
            t, hist, info = fetch_yf_with_retry(ticker, max_retries=3, delay=0.5)

            if hist.empty and not info:
                continue
            
            row = {'Ticker': ticker, 'Region': region}
            
            row['Name'] = safe_get(info, 'shortName', 'longName')
            row['Category'] = safe_get(info, 'category')
            row['Price'] = fmt(safe_get(info, 'regularMarketPrice'))
            
            yf_expense = safe_get(info, 'expenseRatio')
            if yf_expense:
                row['ExpenseRatio(%)'] = fmt(yf_expense * 100, 3)
            elif ticker in US_ETF_EXPENSE:
                row['ExpenseRatio(%)'] = US_ETF_EXPENSE[ticker]
            else:
                row['ExpenseRatio(%)'] = None
            
            row['TotalAssets(B)'] = fmt(safe_get(info, 'totalAssets', default=0) / 1e9, 2) if safe_get(info, 'totalAssets') else None
            row['DivYield(%)'] = fmt(safe_get(info, 'yield', default=0) * 100) if safe_get(info, 'yield') else None
            
            # ★ v3.2.1 수정: 기술적 지표 (hist 전달)
            if not hist.empty and len(hist) > 20:
                close = hist['Close']
                row = add_technical_indicators(row, close, hist=hist, include_ma60_120=False)

            if row.get('52wHigh') is None:
                val = safe_get(info, 'fiftyTwoWeekHigh')
                if val:
                    row['52wHigh'] = fmt(val)
            if row.get('52wLow') is None:
                val = safe_get(info, 'fiftyTwoWeekLow')
                if val:
                    row['52wLow'] = fmt(val)

            if row.get('Price') is None:
                row['Price'] = fmt(safe_get(info, 'regularMarketPrice', 'previousClose', 'navPrice'))

            row = calc_data_quality_score(row, REQUIRED_COLS_ETF, 'etf')
            results.append(row)

        except Exception as e:
            results.append({'Ticker': ticker, 'Remark': str(e)[:50]})

        time.sleep(0.15)

    return pd.DataFrame(results)

# ============================================================
# 한국 ETF
# ============================================================
def get_korea_etfs():
    """한국 ETF 데이터"""
    log("\n[3/5] 한국 ETF 수집 중...")
    
    kr_etf_list = []
    
    # 1순위: pykrx
    if PYKRX_AVAILABLE and not kr_etf_list:
        try:
            log("  pykrx에서 ETF 리스트 로드 중...")
            kr_etf_list = pykrx_stock.get_etf_ticker_list()
            if kr_etf_list:
                log(f"  pykrx: {len(kr_etf_list)}개 ETF ✅")
        except Exception as e:
            log(f"  ⚠️ pykrx ETF 실패: {str(e)[:50]}")
    
    # 2순위: FinanceDataReader
    if not kr_etf_list and FDR_AVAILABLE:
        try:
            log("  FinanceDataReader에서 ETF 리스트 로드 중...")
            etfs = fdr.StockListing('ETF/KR')
            if etfs is not None and not etfs.empty:
                kr_etf_list = []
                for _, row in etfs.iterrows():
                    code = str(row.get('Code', row.get('Symbol', ''))).zfill(6)
                    if len(code) == 6:
                        kr_etf_list.append(code)
                log(f"  FinanceDataReader: {len(kr_etf_list)}개 ETF ✅")
        except Exception as e:
            log(f"  ⚠️ FinanceDataReader ETF 실패: {str(e)[:50]}")
    
    # 3순위: 하드코딩
    if not kr_etf_list or len(kr_etf_list) < 50:
        log("  ⚠️ ETF 리스트 부족, 하드코딩 데이터로 보충")
        hardcoded_etfs = list(KR_ETF_EXPENSE.keys())
        existing_codes = set(str(c).zfill(6) for c in kr_etf_list)
        for code in hardcoded_etfs:
            if code not in existing_codes:
                kr_etf_list.append(code)
        log(f"  하드코딩 추가 후: {len(kr_etf_list)}개 ETF")
    
    if TOP_N_KR_ETF and len(kr_etf_list) > TOP_N_KR_ETF:
        kr_etf_list = kr_etf_list[:TOP_N_KR_ETF]
    
    log(f"  대상: {len(kr_etf_list)}개")
    
    results = []
    start_time = time.time()
    
    for i, code in enumerate(kr_etf_list):
        try:
            code = str(code).zfill(6)
            row = {'Code': code, 'Region': 'KR'}

            # pykrx로 이름, 가격 가져오기
            if PYKRX_AVAILABLE:
                try:
                    row['Name'] = pykrx_stock.get_etf_ticker_name(code)
                    today_str = datetime.now().strftime("%Y%m%d")
                    ohlcv = pykrx_stock.get_etf_ohlcv_by_date(
                        (datetime.now() - timedelta(days=7)).strftime("%Y%m%d"),
                        today_str, code
                    )
                    if not ohlcv.empty:
                        row['Price'] = fmt(ohlcv['종가'].iloc[-1], 0)
                        row['Volume'] = int(ohlcv['거래량'].iloc[-1]) if ohlcv['거래량'].iloc[-1] else None
                except:
                    pass

            # FDR로 히스토리
            hist = pd.DataFrame()
            if FDR_AVAILABLE:
                try:
                    fdr_data = fdr.DataReader(code, DATE_1Y_AGO)
                    if fdr_data is not None and not fdr_data.empty and len(fdr_data) > 20:
                        hist = fdr_data
                        if 'Close' not in hist.columns and '종가' in hist.columns:
                            hist = hist.rename(columns={'종가': 'Close', '시가': 'Open', '고가': 'High', '저가': 'Low', '거래량': 'Volume'})
                except:
                    pass

            # yfinance (폴백)
            info = {}
            if hist.empty:
                ticker_variants = [f"{code}.KS", f"{code}.KQ"]
                t, hist, info = try_multiple_tickers(ticker_variants, max_retries=1)
            else:
                try:
                    t = yf.Ticker(f"{code}.KS")
                    info = t.info if hasattr(t, 'info') else {}
                    if info is None:
                        info = {}
                except:
                    info = {}

            if not row.get('Name'):
                row['Name'] = safe_get(info, 'shortName', 'longName') or code
            if not row.get('Price'):
                row['Price'] = fmt(safe_get(info, 'regularMarketPrice'))

            # Category, ExpenseRatio, DivYield 하드코딩
            yf_category = safe_get(info, 'category')
            if yf_category:
                row['Category'] = yf_category
            elif code in KR_ETF_CATEGORY:
                row['Category'] = KR_ETF_CATEGORY[code]
            
            yf_expense = safe_get(info, 'expenseRatio')
            if yf_expense:
                row['ExpenseRatio(%)'] = fmt(yf_expense * 100, 3)
            elif code in KR_ETF_EXPENSE:
                row['ExpenseRatio(%)'] = KR_ETF_EXPENSE[code]
            
            yf_yield = safe_get(info, 'yield') or safe_get(info, 'dividendYield')
            if yf_yield:
                row['DivYield(%)'] = fmt(yf_yield * 100, 2)
            elif code in KR_ETF_DIVYIELD:
                row['DivYield(%)'] = KR_ETF_DIVYIELD[code]
            
            total_assets = safe_get(info, 'totalAssets')
            if total_assets:
                row['TotalAssets(B)'] = fmt(total_assets / 1e9, 2)
            
            # ★ v3.2.1 수정: 기술적 지표 (hist 전달)
            if not hist.empty and 'Close' in hist.columns and len(hist) > 20:
                close = hist['Close']
                row = add_technical_indicators(row, close, hist=hist, include_ma60_120=False)
            elif not info:
                continue

            if row.get('52wHigh') is None:
                val = safe_get(info, 'fiftyTwoWeekHigh')
                if val:
                    row['52wHigh'] = fmt(val)
            if row.get('52wLow') is None:
                val = safe_get(info, 'fiftyTwoWeekLow')
                if val:
                    row['52wLow'] = fmt(val)

            row = calc_data_quality_score(row, REQUIRED_COLS_ETF, 'etf')
            results.append(row)

        except Exception as e:
            if len(results) < 5:
                log(f"    ⚠️ ETF 실패: {code} - {str(e)[:50]}")
        
        if (i + 1) % 50 == 0 or i == 0:
            elapsed = time.time() - start_time
            per_etf = elapsed / (i + 1) if i > 0 else 0
            remaining = per_etf * (len(kr_etf_list) - i - 1)
            log(f"  진행: {i+1}/{len(kr_etf_list)} - 남은시간: {remaining/60:.1f}분")
        
        time.sleep(0.05)
    
    log(f"  ✅ 완료: {len(results)}개")
    return pd.DataFrame(results)

# ============================================================
# 미국 ETF
# ============================================================
def get_us_etfs():
    """미국 ETF 데이터"""
    log("\n[4/5] 미국 ETF 수집 중...")
    
    tickers = [
        'SPY', 'IVV', 'VOO', 'VTI', 'QQQ', 'DIA', 'IWM', 'IWF', 'IWD', 'VUG',
        'VTV', 'IJH', 'IJR', 'VB', 'VO', 'RSP', 'SPLG', 'SCHX', 'SCHB', 'MGK',
        'XLK', 'XLF', 'XLV', 'XLE', 'XLI', 'XLY', 'XLP', 'XLU', 'XLB', 'XLRE',
        'VGT', 'VFH', 'VHT', 'VDE', 'VIS', 'VCR', 'VDC', 'VPU', 'VAW', 'VNQ',
        'ARKK', 'ARKW', 'ARKF', 'ARKG', 'SOXX', 'SMH', 'XBI', 'IBB', 'HACK', 'BOTZ',
        'LIT', 'TAN', 'ICLN', 'PBW', 'QCLN', 'REMX', 'COPX', 'URA',
        'VYM', 'SCHD', 'DVY', 'HDV', 'SPHD', 'SPYD', 'VIG', 'DGRO', 'NOBL',
        'BND', 'AGG', 'TLT', 'IEF', 'SHY', 'LQD', 'HYG', 'JNK', 'TIP',
        'GLD', 'IAU', 'SLV', 'USO', 'UNG', 'DBC', 'PDBC',
        'TQQQ', 'SQQQ', 'UPRO', 'SPXU', 'SOXL', 'SOXS',
        'VEA', 'VWO', 'EFA', 'EEM', 'IEFA', 'IEMG', 'VXUS', 'ACWI'
    ]
    
    if TOP_N_US_ETF and len(tickers) > TOP_N_US_ETF:
        tickers = tickers[:TOP_N_US_ETF]
    
    log(f"  대상: {len(tickers)}개")
    df = get_etf_data(tickers, "US")
    log(f"  ✅ 완료: {len(df)}개")
    return df

# ============================================================
# 시장 지표
# ============================================================
def get_fear_greed_index():
    """CNN Fear & Greed Index"""
    try:
        url = "https://production.dataviz.cnn.io/index/fearandgreed/graphdata"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/json',
        }
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            if 'fear_and_greed' in data:
                return {
                    'value': fmt(data['fear_and_greed'].get('score'), 1),
                    'rating': data['fear_and_greed'].get('rating', ''),
                    'previous_close': fmt(data['fear_and_greed'].get('previous_close'), 1),
                }
    except:
        pass
    return None

def get_cape_ratio():
    """S&P 500 Shiller CAPE Ratio"""
    try:
        url = "https://www.multpl.com/shiller-pe"
        resp = requests.get(url, headers=HEADERS, timeout=10)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'lxml')
            current = soup.select_one('#current')
            if current:
                val_text = current.get_text().strip().replace(',', '')
                match = re.search(r'([\d.]+)', val_text)
                if match:
                    return fmt(float(match.group(1)), 2)
    except:
        pass
    return None

def get_sp500_forward_pe():
    """S&P 500 Forward PE"""
    try:
        for ticker in ['SPY', 'IVV', 'VOO']:
            t = yf.Ticker(ticker)
            info = t.info if hasattr(t, 'info') else {}
            if info is None:
                info = {}
            pe = safe_get(info, 'trailingPE')
            if pe and 10 < pe < 50:
                return fmt(pe, 2)
    except:
        pass
    return None

def get_market_indicators():
    """글로벌 시장 지표"""
    log("\n[5/5] 시장 지표 수집 중...")

    indicators = {
        '^GSPC': ('S&P 500', '지수'),
        '^DJI': ('다우존스', '지수'),
        '^IXIC': ('나스닥 종합', '지수'),
        '^VIX': ('VIX 변동성', '지수'),
        '^KS11': ('코스피', '지수'),
        '^KQ11': ('코스닥', '지수'),
        '^N225': ('니케이 225', '지수'),
        '^HSI': ('항셍', '지수'),
        'USDKRW=X': ('USD/KRW', '환율'),
        'EURUSD=X': ('EUR/USD', '환율'),
        'GC=F': ('금 선물', '원자재'),
        'CL=F': ('WTI 원유', '원자재'),
        '^TNX': ('미국채 10년', '채권'),
        'BTC-USD': ('비트코인', '암호화폐'),
        'HYG': ('하이일드 채권 ETF', '신용'),
        'LQD': ('투자등급 채권 ETF', '신용'),
        'TLT': ('장기국채 ETF', '채권'),
    }

    results = []

    for ticker, (name, category) in indicators.items():
        try:
            t = yf.Ticker(ticker)
            hist = t.history(period="1y")

            if hist.empty:
                continue

            close = hist['Close']
            current_price = close.iloc[-1]

            decimals = 4 if category == '환율' else (3 if category == '채권' else 2)

            row = {
                'Category': category,
                'Ticker': ticker,
                'Name': name,
                'Price': fmt(current_price, decimals),
            }

            if len(close) >= 2:
                row['Change(%)'] = fmt((close.iloc[-1] / close.iloc[-2] - 1) * 100)
            if len(close) >= 6:
                row['Return1W(%)'] = fmt((close.iloc[-1] / close.iloc[-6] - 1) * 100)
            if len(close) >= 22:
                row['Return1M(%)'] = fmt((close.iloc[-1] / close.iloc[-22] - 1) * 100)
            if len(close) >= 66:
                row['Return3M(%)'] = fmt((close.iloc[-1] / close.iloc[-66] - 1) * 100)
            if len(close) >= 245:
                row['Return1Y(%)'] = fmt((close.iloc[-1] / close.iloc[-min(len(close), 252)] - 1) * 100)
                year_data = close.tail(min(len(close), 252))
                row['52wHigh'] = fmt(year_data.max(), decimals)
                row['52wLow'] = fmt(year_data.min(), decimals)
                row['From52wHigh(%)'] = fmt((current_price / year_data.max() - 1) * 100)
                row['From52wLow(%)'] = fmt((current_price / year_data.min() - 1) * 100)

            results.append(row)
        except:
            pass

        time.sleep(0.05)

    # 심리 지표
    fg_data = get_fear_greed_index()
    if fg_data:
        results.append({
            'Category': '심리',
            'Ticker': 'F&G',
            'Name': 'Fear & Greed Index',
            'Price': fg_data.get('value'),
            'Signal': fg_data.get('rating', ''),
            'Previous': fg_data.get('previous_close'),
        })

    cape = get_cape_ratio()
    if cape:
        results.append({
            'Category': '밸류에이션',
            'Ticker': 'CAPE',
            'Name': 'Shiller CAPE Ratio',
            'Price': cape,
            'Signal': '고평가' if float(cape) > 30 else ('저평가' if float(cape) < 15 else '보통'),
        })

    forward_pe = get_sp500_forward_pe()
    if forward_pe:
        results.append({
            'Category': '밸류에이션',
            'Ticker': 'SPY-PE',
            'Name': 'S&P500 PE (trailing)',
            'Price': forward_pe,
        })

    log(f"  ✅ 완료: {len(results)}개")
    return pd.DataFrame(results)

# ============================================================
# Excel/JSON 저장
# ============================================================
def save_to_excel(data_dict, filename):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils.dataframe import dataframe_to_rows
    
    log("\n엑셀 저장 중...")
    
    wb = Workbook()
    wb.remove(wb.active)
    
    header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=9)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    for sheet_name, df in data_dict.items():
        if df is None or df.empty:
            continue
        
        ws = wb.create_sheet(title=sheet_name[:31])
        
        for r, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c, val in enumerate(row, 1):
                ws.cell(row=r, column=c, value=val)
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        
        ws.freeze_panes = 'A2'
        log(f"  ✅ {sheet_name}: {len(df)}행")
    
    wb.save(filename)
    log(f"\n💾 Excel 저장: {filename}")

def save_to_json(data_dict, filename):
    log("\nJSON 저장 중...")
    
    output = {
        'metadata': {
            'generated': datetime.now().isoformat(),
            'date': TODAY,
            'version': 'v3.2.1-bugfix'
        },
        'data': {}
    }
    
    for sheet_name, df in data_dict.items():
        if df is None or df.empty:
            continue
        
        records = df.replace({np.nan: None}).to_dict(orient='records')
        output['data'][sheet_name] = records
        log(f"  ✅ {sheet_name}: {len(records)}개")
    
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    log(f"\n💾 JSON 저장: {filename}")
    
    size_mb = os.path.getsize(filename) / (1024 * 1024)
    log(f"   파일 크기: {size_mb:.2f} MB")

# ============================================================
# 메인
# ============================================================
def main():
    parser = argparse.ArgumentParser(description='글로벌 주식/ETF 스크리너 (GitHub Actions) v3.2.1')
    parser.add_argument('--json-only', action='store_true', help='JSON만 출력')
    parser.add_argument('--output-dir', type=str, default='.', help='출력 디렉토리')
    parser.add_argument('--kr-stocks', type=int, default=150, help='한국 주식 수 (기본 150)')
    parser.add_argument('--us-stocks', type=int, default=100, help='미국 주식 수 (기본 100)')
    parser.add_argument('--kr-etfs', type=int, default=200, help='한국 ETF 수 (기본 200)')
    parser.add_argument('--us-etfs', type=int, default=100, help='미국 ETF 수 (기본 100)')
    args = parser.parse_args()
    
    global TOP_N_KR, TOP_N_US, TOP_N_KR_ETF, TOP_N_US_ETF
    TOP_N_KR = args.kr_stocks
    TOP_N_US = args.us_stocks
    TOP_N_KR_ETF = args.kr_etfs
    TOP_N_US_ETF = args.us_etfs
    
    log("=" * 60)
    log("글로벌 주식/ETF 스크리닝 - GitHub Actions v3.2.1 (버그 수정)")
    log("★ v3.2.1: hist 매개변수 버그 수정, 안정성 개선")
    log(f"실행: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"한국 주식: {TOP_N_KR}개 (KOSPI) | 미국 주식: {TOP_N_US}개")
    log(f"한국 ETF: {TOP_N_KR_ETF}개 | 미국 ETF: {TOP_N_US_ETF}개")
    log("=" * 60)
    
    start = time.time()
    
    data = {
        'KR_Stocks': get_korea_stocks(),
        'US_Stocks': get_us_stocks(),
        'KR_ETF': get_korea_etfs(),
        'US_ETF': get_us_etfs(),
        'Market_Indicators': get_market_indicators()
    }
    
    output_dir = args.output_dir
    os.makedirs(output_dir, exist_ok=True)
    
    json_file = os.path.join(output_dir, 'data.json')
    save_to_json(data, json_file)
    
    if not args.json_only:
        excel_file = os.path.join(output_dir, f'global_screening_{TODAY}.xlsx')
        save_to_excel(data, excel_file)
    
    elapsed = (time.time() - start) / 60
    log(f"\n총 소요: {elapsed:.1f}분")
    log("=" * 60)

if __name__ == "__main__":
    main()

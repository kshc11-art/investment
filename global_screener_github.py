#!/usr/bin/env python3
"""
글로벌 주식/ETF 스크리닝 - GitHub Actions 버전 v3.2.0
=============================================================================
v3.2.0 신규:
  1. 종목 리스트 통합 수집 (pykrx + KRX + FDR + 네이버 + NXT 합치기)
  2. 중복 제거: 코드(숫자 6자리) 기준으로 통합
  3. 시세 데이터: NXT > 네이버 > pykrx > yfinance 순차 적용
  4. KST 08:00 실행 (NXT 프리마켓 직후 최신 데이터)
=============================================================================
v3.0.2 신규:
  1. FinanceDataReader 추가
  2. 한국 주식/ETF 하드코딩 폴백 강화
=============================================================================
v3.0.1 버그 수정:
  1. Return1Y 인덱스 버그 수정 (min(len-1, 252) → min(len, 252))
  2. 무위험수익률 통일 (2% → 4%)
  3. KR_ETF TotalAssets 단위 수정 (1e12 → 1e9)
=============================================================================
v3.0 수정사항 (데이터 누락 해결):
  1. SharpeRatio: 계산 조건 완화 (252일 → 200일)
  2. Return1Y/Return250D: 계산 조건 완화 (252일 → 245일)
  3. US_Stocks: MA60, MA120, vs_MA60(%), vs_MA120(%) 추가
  4. US_ETF: ExpenseRatio 하드코딩 (yfinance 미제공 대응)
  5. KR_ETF: ExpenseRatio, DivYield, Category 하드코딩
  6. 모든 시트: SharpeRatio 계산 로직 수정
=============================================================================

GitHub Actions에서 자동 실행 → JSON 출력 → GitHub Pages에서 PWA가 fetch

설치:
pip install yfinance openpyxl pandas requests beautifulsoup4 lxml numpy pykrx finance-datareader

실행:
python global_screener_v3.py              # Excel + JSON 출력
python global_screener_v3.py --json-only  # JSON만 출력
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

# KR ETF 비용비율 (총보수, %) - yfinance가 제공하지 않음
KR_ETF_EXPENSE = {
    '069500': 0.05,   # KODEX 200
    '114800': 0.07,   # KODEX 인버스
    '122630': 0.07,   # KODEX 레버리지
    '229200': 0.25,   # KODEX 코스닥150
    '252670': 0.64,   # KODEX 200선물인버스2X
    '305720': 0.45,   # KODEX 2차전지산업
    '091160': 0.45,   # KODEX 반도체
    '133690': 0.07,   # TIGER 나스닥100
    '143850': 0.07,   # TIGER S&P500
    '192090': 0.55,   # TIGER 차이나CSI300
    '371460': 0.07,   # TIGER 미국나스닥100커버드콜
    '360200': 0.25,   # KINDEX 미국다우존스
    '148020': 0.45,   # KODEX MSCI KOREA
    '153130': 0.24,   # KODEX 단기채권
    '161510': 0.04,   # ARIRANG 고배당주
    '157450': 0.04,   # TIGER 고배당
    '261240': 0.35,   # KODEX USD Futures
    '360750': 0.07,   # TIGER S&P500
    '371450': 0.49,   # TIGER 글로벌클라우드컴퓨팅
    '379800': 0.07,   # KODEX S&P500
    '381170': 0.49,   # TIGER 미국테크TOP10 INDXX
    '395160': 0.45,   # KODEX AI반도체핵심장비
    '461460': 0.10,   # PLUS 10년국채액티브
    '102110': 0.05,   # TIGER 200
    '105190': 0.07,   # KINDEX 200
    '226490': 0.07,   # KODEX 코스피
    '102780': 0.07,   # KODEX 삼성그룹
    '091170': 0.45,   # KODEX 은행
    '091180': 0.45,   # KODEX 자동차
    '117680': 0.45,   # KODEX 건설
    '117700': 0.45,   # KODEX 철강
    '139260': 0.45,   # TIGER 미디어컨텐츠
    '139280': 0.45,   # TIGER 경기방어
    '140700': 0.07,   # KODEX 국채선물
    '148070': 0.50,   # KOSEF 단기자금
    '152500': 0.07,   # KINDEX 레버리지
    '156080': 0.50,   # KODEX MSCI World
    '167860': 0.45,   # KOSEF 미국달러선물
    '182490': 0.07,   # TIGER 나스닥바이오
    '195930': 0.45,   # TIGER 유로스탁스50
    '195980': 0.50,   # ARIRANG 신흥국MSCI
    '200250': 0.07,   # KOSEF 미국S&P500선물
    '210780': 0.07,   # TIGER 코스피고배당
    '211210': 0.07,   # KODEX 종합채권
    '214980': 0.07,   # KODEX 단기채권PLUS
    '217770': 0.50,   # TIGER 원유선물
    '219390': 0.07,   # KINDEX 미국S&P500
    '219480': 0.50,   # KODEX 미국S&P500선물
    '226380': 0.50,   # KINDEX 유로스탁스50
    '227830': 0.07,   # ARIRANG 코스피50
    '228790': 0.07,   # TIGER 화장품
    '228800': 0.07,   # TIGER 여행레저
    '228810': 0.07,   # TIGER 미디어컨텐츠
    '228820': 0.07,   # TIGER 은행
    '233740': 0.50,   # TIGER 원유선물Enhanced
    '233160': 0.50,   # TIGER 코스닥150레버리지
    '238720': 0.07,   # KINDEX 코스닥150
    '243880': 0.07,   # TIGER 200에너지화학
    '243890': 0.07,   # TIGER 200건설
    '244580': 0.50,   # KODEX 골드선물
    '244620': 0.50,   # KODEX 은선물
    '245340': 0.50,   # TIGER 리츠부동산인프라
    '245710': 0.07,   # TIGER 코스닥150
    '251340': 0.50,   # KODEX 선진국MSCI World
    '252400': 0.50,   # KODEX 200동일가중
    '252710': 0.07,   # TIGER 200선물레버리지
    '253150': 0.50,   # ARIRANG 고배당저변동
    '253160': 0.50,   # ARIRANG 스마트베타Quality
    '253240': 0.50,   # KODEX 코스닥150선물인버스
    '261110': 0.07,   # TIGER 코스피대형주
    '261120': 0.07,   # TIGER 코스피중형주
    '266370': 0.50,   # KODEX 인도Nifty50
    '267440': 0.07,   # TIGER 코스닥150IT
    '267490': 0.50,   # KINDEX 인버스
    '267500': 0.50,   # KINDEX 레버리지
    '268280': 0.50,   # KINDEX 200선물레버리지
    '269420': 0.07,   # KODEX S&P글로벌인프라
    '270800': 0.50,   # TIGER 미국채10년선물
    '272220': 0.50,   # TIGER 코스닥150선물인버스
    '272560': 0.50,   # TIGER S&P500선물인버스
    '272580': 0.07,   # TIGER 미국S&P500
    '273130': 0.07,   # KODEX 종합채권(AA-)
    '273140': 0.07,   # KODEX 현금배당성장
    '273210': 0.50,   # ARIRANG 미국S&P500
    '273220': 0.50,   # ARIRANG 미국나스닥100
    '275980': 0.07,   # TIGER 200커버드콜5%OTM
    '276650': 0.07,   # TIGER 은행고배당
    '276990': 0.45,   # KINDEX 일본TOPIX100
    '277630': 0.45,   # TIGER 코스피대형가치
    '277640': 0.45,   # TIGER 코스피대형성장
    '278240': 0.45,   # ARIRANG 스마트베타Momentum
    '278420': 0.07,   # KODEX 코스피대형주
    '278530': 0.45,   # TIGER 200에너지화학레버리지
    '278540': 0.45,   # TIGER 200IT레버리지
    '284430': 0.07,   # KODEX 코스닥150선물레버리지
    '284980': 0.10,   # HANARO 200
    '287300': 0.10,   # HANARO KOSDAQ150
    '287310': 0.07,   # KINDEX 코스피
    '287320': 0.50,   # KINDEX 미국달러선물레버리지
    '287330': 0.50,   # KINDEX 미국달러선물인버스
    '290080': 0.50,   # KODEX 미국채Ultra30년선물
    '292150': 0.50,   # TIGER 미국채30년선물
    '292160': 0.50,   # TIGER 미국채10년선물레버리지
    '292170': 0.50,   # TIGER 미국채10년선물인버스
    '292180': 0.50,   # TIGER 미국채10년선물인버스2X
    '292190': 0.50,   # TIGER 미국채30년선물인버스
    '294400': 0.50,   # KODEX 미국채Ultra10년선물
    '298340': 0.45,   # TIGER 2차전지테마
    '298770': 0.45,   # KODEX 미국채10년선물
    '299660': 0.07,   # TIGER 200산업재
    '300640': 0.50,   # KODEX 코스닥150선물레버리지
    '300950': 0.50,   # TIGER 나스닥바이오텍
    '302190': 0.50,   # TIGER 미국달러선물레버리지
    '304770': 0.45,   # TIGER 게임
    '304780': 0.45,   # TIGER 바이오
    '304940': 0.45,   # KODEX 바이오
    '305080': 0.45,   # TIGER 미디어콘텐츠
    '307510': 0.50,   # ARIRANG ESG종합채권
    '315960': 0.50,   # ARIRANG KS채권혼합
    '319870': 0.50,   # TIGER 200커버드콜ATM
    '322120': 0.50,   # TIGER 통신TOP10
    '322130': 0.45,   # TIGER 의료기기
    '322400': 0.50,   # KODEX 테슬라인컴인버스
    '322410': 0.50,   # KODEX 테슬라인컴레버리지
    '322500': 0.50,   # KODEX 테슬라인컴
    '329750': 0.50,   # TIGER 미국테크TOP10INDXX
    '333940': 0.50,   # ARIRANG S&P500
    '337140': 0.50,   # KODEX 3대농산물선물
    '337160': 0.45,   # KODEX 게임산업
    '352540': 0.45,   # TIGER 헬스케어
    '352560': 0.50,   # TIGER K리츠
    '354350': 0.50,   # KODEX 인버스2X
    '360140': 0.50,   # TIGER AI코리아그로스
    '363570': 0.50,   # KODEX 자동차
    '363580': 0.50,   # KODEX 은행
    '364970': 0.50,   # TIGER KRX바이오K뉴딜
    '364980': 0.50,   # TIGER KRX2차전지K뉴딜
    '365000': 0.50,   # TIGER KRX인터넷K뉴딜
    '365040': 0.50,   # KODEX K뉴딜디지털플러스
    '367380': 0.50,   # TIGER 미국AI테크TOP10
    '368190': 0.50,   # TIGER AI&로봇
    '368590': 0.50,   # KODEX 글로벌리튬
    '371150': 0.50,   # TIGER KRX BBIG K뉴딜
    '371160': 0.50,   # TIGER 차이나전기차레버리지
    '372790': 0.50,   # TIGER 반도체TOP10
    '373490': 0.50,   # TIGER 우주방산
    '375270': 0.50,   # TIGER 코스닥150리밸런싱
    '375720': 0.50,   # TIGER 2차전지테크
    '381180': 0.50,   # TIGER 미국필라델피아반도체
    '385720': 0.50,   # TIGER Fn반도체TOP10
    '385550': 0.50,   # KODEX K-방산
    '385560': 0.50,   # KODEX 글로벌K뉴딜
    '385590': 0.50,   # TIGER K로봇
    '385600': 0.50,   # TIGER K AI반도체핵심장비
    '387270': 0.50,   # TIGER 차이나항셍테크레버리지
    '387280': 0.50,   # TIGER 차이나항셍테크인버스
    '391600': 0.50,   # TIGER AI반도체핵심소재
    '391680': 0.50,   # TIGER AI코리아펀더멘탈
    '394660': 0.50,   # TIGER 글로벌자율주행
    '394670': 0.50,   # TIGER 글로벌BBIG핀테크
    '396070': 0.50,   # KODEX K방산
    '400760': 0.50,   # TIGER 글로벌2차전지TOP10
    '401470': 0.50,   # TIGER 글로벌AI로봇&자율주행
    '402340': 0.50,   # HANARO 글로벌AI에너지
    '404780': 0.50,   # HANARO 글로벌탄소중립
    '411060': 0.50,   # KODEX K게임
    '445280': 0.50,   # TIGER AI BIGTECH 10
    '453810': 0.50,   # TIGER 미국테크TOP10커버드콜
    '458730': 0.50,   # TIGER 미국30년국채스트립액티브
    '459000': 0.50,   # TIGER 미국채30년스트립액티브
    '459100': 0.50,   # KODEX 미국30년국채스트립액티브
    '459200': 0.50,   # TIGER 미국10년국채스트립액티브
    '459580': 0.50,   # TIGER 글로벌온디바이스AI
    '462320': 0.50,   # TIGER 미국배당다우존스
    '464510': 0.50,   # TIGER 미국캐시카우100커버드콜
    # ★ v3.0.2 추가: 주요 ETF 보강
    '069660': 0.45,   # KOSEF 200
    '091230': 0.45,   # TIGER 반도체
    '098560': 0.50,   # TIGER 방송통신
    '099140': 0.50,   # KODEX China H
    '100910': 0.50,   # KOSEF 단기자금
    '101280': 0.50,   # KODEX Japan
    '104520': 0.45,   # KOSEF 블루칩
    '108440': 0.50,   # KINDEX Japan
    '108450': 0.50,   # KINDEX China
    '108480': 0.50,   # KINDEX 인도네시아
    '114260': 0.50,   # KODEX 국채3년
    '114470': 0.07,   # KOSEF 국채3년
    '114820': 0.50,   # TIGER 국채3년
    '117460': 0.50,   # KODEX 에너지화학
    '123310': 0.50,   # TIGER 인버스
    '123320': 0.50,   # TIGER 레버리지
    '130680': 0.50,   # TIGER 원자재
    '130730': 0.50,   # KOSEF 원자재
    '131890': 0.50,   # KINDEX 인버스
    '132030': 0.50,   # KODEX 골드선물(H)
    '136340': 0.50,   # KINDEX 밸류대형
    '137610': 0.50,   # TIGER 농산물선물
    '137930': 0.50,   # TIGER 금속선물
    '138230': 0.50,   # KOSEF 미국달러선물
    '138910': 0.50,   # KODEX 구리선물
    '138920': 0.50,   # KODEX 콩선물
    '139220': 0.45,   # TIGER 200 헬스케어
    '139230': 0.45,   # TIGER 200 중공업
    '139240': 0.45,   # TIGER 200 건설
    '139250': 0.45,   # TIGER 200 경기소비재
    '139270': 0.45,   # TIGER 200 금융
    '139290': 0.45,   # TIGER 200 철강소재
    '139310': 0.45,   # TIGER 200 에너지화학
    '139320': 0.45,   # TIGER 200 IT
    '140570': 0.50,   # KINDEX 단기통안채
    '140580': 0.50,   # KINDEX China
    '140710': 0.50,   # KODEX 10년국채선물
    '145850': 0.50,   # TIGER 일본TOPIX
    '147970': 0.50,   # TIGER 모멘텀
    '148040': 0.50,   # KINDEX 200동일가중
    '148060': 0.50,   # KINDEX 배당성장
    '150460': 0.50,   # KINDEX 200선물인버스
    '152100': 0.07,   # ARIRANG 200
    '152180': 0.50,   # TIGER 미국MSCI리츠
    '153270': 0.50,   # KODEX 건설
    '156080': 0.50,   # KODEX MSCI World
    '159800': 0.50,   # 마이다스 200커버드콜
    '160580': 0.50,   # TIGER 구리실물
    '169950': 0.50,   # KODEX China H레버리지
    '176710': 0.50,   # 파워 국채10년
    '176950': 0.50,   # KODEX 금선물(H)
    '181480': 0.50,   # TIGER 배당성장
    '182480': 0.50,   # TIGER US리츠
    '183700': 0.50,   # KINDEX 중국본토대형
    '185680': 0.50,   # KINDEX 일본레버리지
    '189400': 0.50,   # ARIRANG AC월드
    '190150': 0.50,   # ARIRANG MSCI유럽
    '190620': 0.50,   # KINDEX 러시아
    '192720': 0.50,   # 파워 단기채
    '196030': 0.50,   # KODEX China A50
    '196220': 0.50,   # KINDEX 중국본토CSI300
    '196230': 0.50,   # KINDEX 골드선물
    '200020': 0.50,   # KODEX 미국달러선물인버스
    '203780': 0.50,   # TIGER 차이나A300
    '204420': 0.50,   # ARIRANG 차이나H
    '204450': 0.50,   # ARIRANG 차이나본토
    '204480': 0.50,   # TIGER 차이나A레버리지
    '205720': 0.50,   # KINDEX 일본인버스
    '210950': 0.50,   # KINDEX 코스닥150레버리지
    '211560': 0.50,   # TIGER 배당프리미엄
    '213630': 0.50,   # KODEX 삼성그룹밸류
    '214420': 0.50,   # KODEX China A50인버스
    '215620': 0.50,   # TIGER S&P500동일가중
    '217480': 0.50,   # KINDEX 코스피레버리지
    '218420': 0.50,   # KODEX 미국채10년선물
    '219390': 0.50,   # KINDEX 미국S&P500
    '220130': 0.50,   # SMART 200
    '222180': 0.50,   # ARIRANG 스마트베타
    '222190': 0.50,   # ARIRANG 스마트베타Value
    '222200': 0.50,   # ARIRANG 스마트베타Momentum
    '223190': 0.50,   # KODEX 200중소형
    '224100': 0.50,   # KINDEX 중국본토A300
    '225030': 0.50,   # KINDEX 코스피인버스
    '225040': 0.50,   # KINDEX 코스닥150인버스
    '225050': 0.50,   # KINDEX 코스피
    '225130': 0.50,   # KINDEX 코스닥150
    '226980': 0.50,   # KODEX 200 IT
    '227540': 0.50,   # TIGER 200 헬스케어
    '227550': 0.50,   # TIGER 200 산업재
    '227560': 0.50,   # TIGER 200 생활소비재
    '227570': 0.50,   # TIGER 200 중공업
    '227830': 0.50,   # ARIRANG 코스피50
    '228790': 0.50,   # TIGER 화장품
    '228810': 0.50,   # TIGER 미디어컨텐츠
    '228820': 0.50,   # TIGER 은행
    '232080': 0.50,   # TIGER 코스닥150바이오테크
    '234310': 0.50,   # KODEX 미국메타버스나스닥액티브
    '236350': 0.50,   # TIGER 인도니프티50
    '238670': 0.50,   # KINDEX 미국나스닥100
    '238710': 0.50,   # KINDEX 코스닥150선물인버스
    '240180': 0.50,   # TIGER 로우볼
    '241180': 0.50,   # TIGER 코스닥150헬스케어
    '241390': 0.50,   # KINDEX 미국나스닥100선물
    '241560': 0.50,   # TIGER TOP10
    '245710': 0.50,   # TIGER 코스닥150
    '246710': 0.50,   # KINDEX 미국고배당S&P
    '249580': 0.50,   # KINDEX 코스닥150선물레버리지
    '250730': 0.50,   # KINDEX 한류
    '251590': 0.50,   # ARIRANG 배당주채권혼합
    '261140': 0.50,   # TIGER 코스피중형주
    '261270': 0.50,   # TIGER 코스닥중소형
    '267770': 0.50,   # KODEX 200가치저변동
    '270810': 0.50,   # KINDEX 미국인터넷
    '277630': 0.50,   # TIGER 코스피대형가치
    '277640': 0.50,   # TIGER 코스피대형성장
    '277650': 0.50,   # TIGER 코스피중형가치
    '277660': 0.50,   # TIGER 코스피중형성장
    '278620': 0.50,   # KINDEX 200인버스
    '280920': 0.50,   # KODEX 미국빅테크10
    '280930': 0.50,   # KINDEX 미국빅테크TOP7Plus
    '282330': 0.50,   # KODEX 2차전지산업레버리지
    '283580': 0.50,   # KINDEX 미국채10년선물
    '283590': 0.50,   # KINDEX 미국채10년선물인버스
    '287330': 0.50,   # KINDEX 미국달러선물인버스
    '289480': 0.50,   # KINDEX 신흥국하이일드
    '291890': 0.50,   # KINDEX 블룸버그선진국
    '294020': 0.50,   # KINDEX 미국AI테크
    '295000': 0.50,   # KINDEX 미국친환경그린테마
    '295040': 0.50,   # KINDEX 글로벌클린에너지
    '296900': 0.50,   # KINDEX 미국S&P배당귀족
    '298770': 0.50,   # KODEX 미국채10년선물
    '304660': 0.50,   # KODEX 고배당
    '306520': 0.50,   # KINDEX 중국항셍테크
    '306950': 0.50,   # KINDEX 차이나테크
    '307000': 0.50,   # KINDEX KRX300
    '314250': 0.50,   # KINDEX 미국S&P500인버스
    '315270': 0.50,   # KODEX 글로벌리튬
    '319870': 0.50,   # TIGER 200커버드콜ATM
    '326240': 0.50,   # KINDEX 미국나스닥100(H)
    '329200': 0.50,   # TIGER 미국테크TOP10
    '331910': 0.50,   # KINDEX 글로벌탄소배출권
    '334690': 0.50,   # KINDEX 미국배당귀족나스닥
    '334700': 0.50,   # KINDEX 글로벌메타버스
    '337140': 0.50,   # KODEX 3대농산물선물
    '348580': 0.50,   # KINDEX 미국고배당S&P배당귀족
    '352560': 0.50,   # TIGER K리츠
    '357870': 0.50,   # KINDEX 미국메타버스
    '357880': 0.50,   # KINDEX 미국S&P500ESG
    '359210': 0.50,   # KINDEX 글로벌전기차&배터리
    '360750': 0.10,   # TIGER S&P500
    '361580': 0.50,   # KINDEX 미국S&P500데일리커버드콜
    '361600': 0.50,   # KINDEX 글로벌수소경제
    '363570': 0.50,   # KODEX 자동차
    '364970': 0.50,   # TIGER KRX바이오K뉴딜
    '364980': 0.50,   # TIGER KRX2차전지K뉴딜
    '365040': 0.50,   # KODEX K뉴딜디지털플러스
    '368200': 0.50,   # KINDEX 미국다우존스
    '371150': 0.50,   # TIGER KRX BBIG K뉴딜
    '372790': 0.50,   # TIGER 반도체TOP10
    '375270': 0.50,   # TIGER 코스닥150리밸런싱
    '375720': 0.50,   # TIGER 2차전지테크
    '375770': 0.50,   # KINDEX 중국클린에너지
    '381180': 0.50,   # TIGER 미국필라델피아반도체
    '385590': 0.50,   # TIGER K로봇
    '385720': 0.50,   # TIGER Fn반도체TOP10
    '391600': 0.50,   # TIGER AI반도체핵심소재
    '394660': 0.50,   # TIGER 글로벌자율주행
    '396690': 0.50,   # KINDEX 글로벌혁신블루칩
    '400760': 0.50,   # TIGER 글로벌2차전지TOP10
    '401470': 0.50,   # TIGER 글로벌AI로봇&자율주행
    '411060': 0.50,   # KODEX K게임
    '430570': 0.50,   # ACE 미국S&P500
    '430600': 0.50,   # ACE 미국나스닥100
    '441660': 0.50,   # KINDEX 미국S&P500퀄리티
    '441680': 0.50,   # KINDEX 미국나스닥100퀄리티
    '445280': 0.50,   # TIGER AI BIGTECH 10
    '446720': 0.50,   # SOL 미국S&P500
    '446770': 0.50,   # SOL 미국나스닥100
    '447770': 0.50,   # TIGER 차이나항셍테크
    '448290': 0.50,   # KINDEX 미국빅테크
    '448320': 0.50,   # KINDEX 인도Nifty50
    '448540': 0.50,   # KINDEX 일본TOPIX100
    '448810': 0.50,   # ACE 미국빅테크TOP7Plus
    '449580': 0.50,   # KINDEX 미국AI인프라
    '450710': 0.50,   # KINDEX 미국반도체MV
    '452550': 0.50,   # KINDEX 미국방어
    '455030': 0.50,   # KINDEX 미국S&P500동일가중
}

# KR ETF 카테고리
KR_ETF_CATEGORY = {
    '069500': '국내 대형주',
    '114800': '인버스',
    '122630': '레버리지',
    '229200': '코스닥',
    '252670': '인버스 레버리지',
    '305720': '섹터(2차전지)',
    '091160': '섹터(반도체)',
    '133690': '미국 대형주',
    '143850': '미국 대형주',
    '192090': '중국',
    '371460': '커버드콜',
    '360200': '미국 대형주',
    '148020': '국내 전체',
    '153130': '채권(단기)',
    '161510': '배당',
    '157450': '배당',
    '261240': '통화(달러)',
    '360750': '미국 대형주',
    '371450': '테마(클라우드)',
    '379800': '미국 대형주',
    '381170': '미국 테크',
    '395160': '섹터(AI반도체)',
    '461460': '채권(10년)',
    '102110': '국내 대형주',
    '105190': '국내 대형주',
    '226490': '국내 전체',
    '102780': '섹터(삼성그룹)',
    '091170': '섹터(은행)',
    '091180': '섹터(자동차)',
    '117680': '섹터(건설)',
    '117700': '섹터(철강)',
    '139260': '섹터(미디어)',
    '139280': '섹터(경기방어)',
}

# KR ETF 배당수익률 (2024년 기준 추정)
KR_ETF_DIVYIELD = {
    '069500': 1.8,   # KODEX 200
    '114800': 0.0,   # KODEX 인버스
    '122630': 0.0,   # KODEX 레버리지
    '229200': 0.3,   # KODEX 코스닥150
    '252670': 0.0,   # 인버스2X
    '305720': 0.0,   # 2차전지
    '091160': 0.5,   # 반도체
    '133690': 0.4,   # 나스닥100
    '143850': 1.2,   # S&P500
    '192090': 0.8,   # 차이나
    '161510': 4.5,   # 고배당주
    '157450': 3.8,   # TIGER 고배당
    '261240': 0.0,   # 달러선물
    '360750': 1.2,   # S&P500
    '371450': 0.0,   # 클라우드
    '379800': 1.2,   # S&P500
    '381170': 0.2,   # 테크TOP10
    '395160': 0.0,   # AI반도체
    '461460': 3.0,   # 국채
    '371460': 8.0,   # 커버드콜
    '102110': 1.8,   # TIGER 200
    '105190': 1.8,   # KINDEX 200
}

# ★ v3.0.2 추가: 한국 주식 하드코딩 (KOSPI 시총 상위 200개)
# 형식: {종목코드: (종목명, 섹터)}
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
    '024110': ('기업은행', '은행'),
    '000810': ('삼성화재', '보험'),
    '259960': ('크래프톤', '서비스업'),
    '011200': ('HMM', '운수창고'),
    '010950': ('S-Oil', '화학'),
    '034020': ('두산에너빌리티', '기계'),
    '009540': ('HD한국조선해양', '운수장비'),
    '036570': ('엔씨소프트', '서비스업'),
    '003490': ('대한항공', '운수창고'),
    '138040': ('메리츠금융지주', '기타금융'),
    '047050': ('포스코인터내셔널', '유통업'),
    '302440': ('SK바이오사이언스', '의약품'),
    '352820': ('하이브', '서비스업'),
    '011170': ('롯데케미칼', '화학'),
    '030200': ('KT', '통신업'),
    '267250': ('HD현대', '운수장비'),
    '032640': ('LG유플러스', '통신업'),
    '090430': ('아모레퍼시픽', '화학'),
    '012450': ('한화에어로스페이스', '기계'),
    '011070': ('LG이노텍', '전기전자'),
    '010140': ('삼성중공업', '운수장비'),
    '051900': ('LG생활건강', '화학'),
    '161390': ('한국타이어앤테크놀로지', '화학'),
    '088980': ('맥쿼리인프라', '기타금융'),
    '036460': ('한국가스공사', '전기가스'),
    '329180': ('HD현대중공업', '운수장비'),
    '004020': ('현대제철', '철강금속'),
    '028050': ('삼성엔지니어링', '건설업'),
    '000720': ('현대건설', '건설업'),
    '326030': ('SK바이오팜', '의약품'),
    '004990': ('롯데지주', '유통업'),
    '042660': ('한화오션', '운수장비'),
    '078930': ('GS', '유통업'),
    '006800': ('미래에셋증권', '증권'),
    '021240': ('코웨이', '전기전자'),
    '097950': ('CJ제일제당', '음식료'),
    '034220': ('LG디스플레이', '전기전자'),
    '069500': ('KODEX 200', 'ETF'),  # ETF 제외용
    '241560': ('두산밥캣', '기계'),
    '377300': ('카카오페이', '서비스업'),
    '000100': ('유한양행', '의약품'),
    '180640': ('한진칼', '운수창고'),
    '016360': ('삼성증권', '증권'),
    '001040': ('CJ', '유통업'),
    '071050': ('한국금융지주', '기타금융'),
    '009830': ('한화솔루션', '화학'),
    '005940': ('NH투자증권', '증권'),
    '006260': ('LS', '전기전자'),
    '006360': ('GS건설', '건설업'),
    '002790': ('아모레G', '화학'),
    '000990': ('DB하이텍', '전기전자'),
    '128940': ('한미약품', '의약품'),
    '035250': ('강원랜드', '서비스업'),
    '011780': ('금호석유', '화학'),
    '001570': ('금양', '화학'),
    '047810': ('한국항공우주', '운수장비'),
    '010620': ('현대미포조선', '운수장비'),
    '271560': ('오리온', '음식료'),
    '005830': ('DB손해보험', '보험'),
    '282330': ('BGF리테일', '유통업'),
    '383220': ('F&F', '섬유의복'),
    '139480': ('이마트', '유통업'),
    '003410': ('쌍용C&E', '비금속'),
    '029780': ('삼성카드', '기타금융'),
    '005387': ('현대차2우B', '운수장비'),
    '002380': ('KCC', '화학'),
    '000080': ('하이트진로', '음식료'),
    '361610': ('SK아이이테크놀로지', '화학'),
    '272210': ('한화시스템', '전기전자'),
    '008770': ('호텔신라', '서비스업'),
    '001450': ('현대해상', '보험'),
    '023530': ('롯데쇼핑', '유통업'),
    '011790': ('SKC', '화학'),
    '064350': ('현대로템', '운수장비'),
    '039490': ('키움증권', '증권'),
    '007070': ('GS리테일', '유통업'),
    '024720': ('한국전자금융', '서비스업'),
    '402340': ('SK스퀘어', '기타금융'),
    '006280': ('녹십자', '의약품'),
    '014680': ('한솔케미칼', '화학'),
    '192820': ('코스맥스', '화학'),
    '008930': ('한미사이언스', '의약품'),
    '052690': ('한전기술', '서비스업'),
    '069620': ('대웅제약', '의약품'),
    '004370': ('농심', '음식료'),
    '030000': ('제일기획', '서비스업'),
    '081660': ('휠라홀딩스', '섬유의복'),
    '214370': ('케어젠', '의약품'),
    '003230': ('삼양식품', '음식료'),
    '251270': ('넷마블', '서비스업'),
    '005389': ('현대차3우B', '운수장비'),
    '175330': ('JB금융지주', '기타금융'),
    '018880': ('한온시스템', '운수장비'),
    '145020': ('휴젤', '의약품'),
    '009240': ('한샘', '기타제조'),
    '010120': ('LS ELECTRIC', '전기전자'),
    '004000': ('롯데정밀화학', '화학'),
    '111770': ('영원무역', '섬유의복'),
    '000120': ('CJ대한통운', '운수창고'),
    '012630': ('HDC', '건설업'),
    '001120': ('LX인터내셔널', '유통업'),
    '000150': ('두산', '기계'),
    '020150': ('일진머티리얼즈', '철강금속'),
    '004170': ('신세계', '유통업'),
    '241840': ('에스에프에이', '기계'),
    '005250': ('녹십자홀딩스', '의약품'),
    '267270': ('HD현대건설기계', '기계'),
    '016800': ('퍼시스', '기타제조'),
    '138930': ('BNK금융지주', '기타금융'),
    '194480': ('데브시스터즈', '서비스업'),
    '004800': ('효성', '화학'),
    '008560': ('메리츠증권', '증권'),
    '017800': ('현대엘리베이터', '기계'),
    '071970': ('STX중공업', '기계'),
    '285130': ('SK케미칼', '화학'),
    '009420': ('한올바이오파마', '의약품'),
    '060980': ('한라홀딩스', '운수장비'),
    '000210': ('DL', '화학'),
    '950160': ('코오롱티슈진', '의약품'),
    '003090': ('대웅', '의약품'),
    '000880': ('한화', '화학'),
    '026960': ('동서', '음식료'),
    '014820': ('동원시스템즈', '기계'),
    '011760': ('현대코퍼레이션', '유통업'),
    '002350': ('넥센타이어', '화학'),
    '007310': ('오뚜기', '음식료'),
    '008350': ('남선알미늄', '철강금속'),
    '069960': ('현대백화점', '유통업'),
    '001740': ('SK네트웍스', '유통업'),
    '003240': ('태광산업', '화학'),
    '044820': ('코스맥스비티아이', '화학'),
    '005610': ('SPC삼립', '음식료'),
    '051600': ('한전KPS', '서비스업'),
    '035510': ('신세계인터내셔날', '유통업'),
    '019170': ('신풍제약', '의약품'),
    '192400': ('쿠쿠홀딩스', '전기전자'),
    '010780': ('아이에스동서', '건설업'),
    '005180': ('빙그레', '음식료'),
    '000490': ('대동', '기계'),
    '021050': ('서원인텍', '전기전자'),
    '001800': ('오리온홀딩스', '유통업'),
    '003850': ('보령', '의약품'),
    '001680': ('대상', '음식료'),
    '057050': ('현대홈쇼핑', '유통업'),
    '033240': ('자화전자', '전기전자'),
    '000070': ('삼양홀딩스', '화학'),
    '006650': ('대한유화', '화학'),
    '002960': ('한국쉘석유', '화학'),
    '093370': ('후성', '화학'),
    '000640': ('동아쏘시오홀딩스', '의약품'),
    '003620': ('쌍용자동차', '운수장비'),
    '060250': ('NHN KCP', '서비스업'),
    '003030': ('세아제강', '철강금속'),
    '138490': ('코오롱플라스틱', '화학'),
    '004910': ('조광페인트', '화학'),
    '089590': ('제주항공', '운수창고'),
    '053210': ('스카이라이프', '서비스업'),
    '161890': ('한국콜마', '화학'),
    '003000': ('부광약품', '의약품'),
    '002030': ('아세아', '비금속'),
    '025860': ('남해화학', '화학'),
    '003220': ('대원강업', '운수장비'),
    '000400': ('롯데손해보험', '보험'),
    '000320': ('노루홀딩스', '화학'),
    '012510': ('더존비즈온', '서비스업'),
}

# ============================================================
# 설정
# ============================================================
TODAY = datetime.now().strftime("%Y%m%d")
DATE_1Y_AGO = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")

# 종목 수 설정
TOP_N_KR = 150    # 한국 주식: KOSPI 150개 고정
TOP_N_US = 100    # 미국 주식: 100개
TOP_N_KR_ETF = 200  # 한국 ETF: 200개
TOP_N_US_ETF = 100  # 미국 ETF: 100개

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
}

# 데이터 소스 차단 여부 (런타임에 판단)
NAVER_AVAILABLE = None
FNGUIDE_AVAILABLE = None
NXT_AVAILABLE = None  # ★ v3.2.0: NXT 차단 감지

# 타임아웃 설정 (GitHub Actions용 단축)
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
            hist = t.history(period="1y")  # 2y → 1y로 단축
            info = t.info

            # 유효한 데이터인지 확인
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

        # 점수 계산: hist 길이 + info 키 수
        score = len(hist) + len(info) if info else len(hist)

        if score > best_score:
            best_score = score
            best_result = (t, hist, info)

            # 충분히 좋은 결과면 조기 종료
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
# ★ v3.2.0: NXT(넥스트레이드) 크롤링
# ============================================================
# NXT 세션 (전역)
_NXT_SESSION = None

def get_nxt_session():
    """NXT 세션 획득 (쿠키 필요)"""
    global _NXT_SESSION
    
    if _NXT_SESSION is not None:
        return _NXT_SESSION
    
    try:
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        
        _NXT_SESSION = requests.Session()
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        }
        
        # 페이지 방문하여 세션/쿠키 획득
        resp = _NXT_SESSION.get(
            'https://www.nextrade.co.kr/menu/transactionStatusMain/menuList.do',
            headers=headers,
            verify=False,
            timeout=10
        )
        
        if resp.status_code == 200:
            log("  NXT 세션 획득 ✅")
            return _NXT_SESSION
        else:
            _NXT_SESSION = None
            return None
    except Exception as e:
        log(f"  ⚠️ NXT 세션 획득 실패: {str(e)[:50]}")
        _NXT_SESSION = None
        return None

def fetch_nxt_api(endpoint, params=None, timeout=TIMEOUT_LONG):
    """NXT API 요청 (세션 기반)"""
    global NXT_AVAILABLE
    
    if NXT_AVAILABLE is False:
        return None
    
    session = get_nxt_session()
    if session is None:
        NXT_AVAILABLE = False
        return None
    
    try:
        url = f"https://www.nextrade.co.kr{endpoint}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'X-Requested-With': 'XMLHttpRequest',
            'Referer': 'https://www.nextrade.co.kr/menu/transactionStatusMain/menuList.do',
        }
        
        resp = session.post(url, data=params, headers=headers, timeout=timeout, verify=False)
        
        if resp.status_code == 403 or resp.status_code == 503:
            NXT_AVAILABLE = False
            log(f"  ⚠️ NXT 접근 차단됨 ({resp.status_code})")
            return None
        
        if resp.status_code == 200:
            NXT_AVAILABLE = True
            try:
                return resp.json()
            except:
                return None
        return None
    except Exception as e:
        log(f"  ⚠️ NXT 요청 실패: {str(e)[:50]}")
        return None

def get_nxt_stock_data():
    """NXT에서 주식 거래현황 데이터 가져오기
    
    Returns:
        dict: {종목코드: {name, price, change, change_rate, volume, market, ...}}
    """
    global NXT_AVAILABLE
    
    if NXT_AVAILABLE is False:
        return {}
    
    log("  NXT 거래현황 로드 시도...")
    
    result = {}
    
    # NXT API - 정규시장 거래현황
    # API가 페이지당 5개씩만 반환 (rows 파라미터 무시)
    # 상위 200개면 충분 → 40페이지
    max_pages = 40
    target_count = 200
    
    for page in range(1, max_pages + 1):
        params = {
            'page': page,
            'rows': 100,  # 무시되지만 일단 전송
            '_search': 'false',
            'sidx': '',
            'sord': 'asc',
        }
        
        data = fetch_nxt_api('/brdinfoTime/brdinfoTimeList.do', params)
        
        if not data:
            if page == 1:
                log("  ⚠️ NXT 데이터 로드 실패")
                NXT_AVAILABLE = False
            break
        
        rows = data.get('brdinfoTimeList', [])
        if not rows:
            break
        
        for row in rows:
            try:
                # NXT API 응답 필드 매핑 (실제 필드명)
                code_raw = str(row.get('isuSrdCd', '')).strip()  # A005930 형태
                code = code_raw.replace('A', '').zfill(6)  # 005930으로 변환
                
                if not code or len(code) != 6:
                    continue
                
                name = row.get('isuAbwdNm', '')  # 종목명 (약어)
                price = row.get('curPrc', 0)  # 현재가
                change = row.get('contrastPrc', 0)  # 전일대비
                change_rate = row.get('upDownRate', 0)  # 등락률
                volume = row.get('accTdQty', 0)  # 누적거래량
                value = row.get('accTrval', 0)  # 누적거래대금
                market = row.get('mktNm', 'KOSPI')  # 시장
                high = row.get('hgpr', 0)  # 고가
                low = row.get('lwpr', 0)  # 저가
                open_price = row.get('oppr', 0)  # 시가
                base_price = row.get('basePrc', 0)  # 기준가
                
                result[code] = {
                    'name': name,
                    'price': int(price) if price else None,
                    'change': int(change) if change else 0,
                    'change_rate': float(change_rate) if change_rate else 0,
                    'volume': int(volume) if volume else 0,
                    'value': int(value) if value else 0,
                    'high': int(high) if high else None,
                    'low': int(low) if low else None,
                    'open': int(open_price) if open_price else None,
                    'base_price': int(base_price) if base_price else None,
                    'market': 'KOSPI' if 'KOSPI' in str(market).upper() else 'KOSDAQ',
                    'source': 'NXT'
                }
            except:
                continue
        
        # 진행 상황 출력 (20페이지마다)
        if page % 20 == 0:
            log(f"    NXT 로딩 중... {len(result)}개")
        
        # 200개 도달 시 종료
        if len(result) >= target_count:
            break
        
        # 마지막 페이지 확인
        total_cnt = data.get('totalCnt', 0)
        if len(result) >= total_cnt:
            break
        
        time.sleep(0.05)  # 속도 조절
    
    if result:
        log(f"  NXT: {len(result)}개 종목 로드 ✅")
        NXT_AVAILABLE = True
    else:
        log("  ⚠️ NXT 데이터 없음")
    
    return result

def get_nxt_etf_data():
    """NXT에서 ETF 거래현황 데이터 가져오기
    (현재 NXT는 ETF 미지원 - 향후 확장 대비)
    """
    # NXT는 2025년 현재 ETF/ETN 미지원
    # 향후 지원 시 이 함수 구현
    return {}

# NXT 데이터 캐시 (전역)
_NXT_STOCK_CACHE = None
_NXT_CACHE_TIME = None

def get_nxt_cached_data():
    """NXT 데이터를 캐시하여 반환 (중복 호출 방지)"""
    global _NXT_STOCK_CACHE, _NXT_CACHE_TIME
    
    # 5분 이내 캐시 사용
    if _NXT_STOCK_CACHE is not None and _NXT_CACHE_TIME is not None:
        if (datetime.now() - _NXT_CACHE_TIME).seconds < 300:
            return _NXT_STOCK_CACHE
    
    _NXT_STOCK_CACHE = get_nxt_stock_data()
    _NXT_CACHE_TIME = datetime.now()
    return _NXT_STOCK_CACHE

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
    
    log(f"  네이버 {market} 시총 순위 로드 시도...")
    
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
                
                change_pct = None
                if len(tds) > 4:
                    try:
                        pct_text = tds[4].text.strip().replace('%', '').replace('+', '')
                        change_pct = float(pct_text) if pct_text else None
                    except:
                        pass
                
                market_cap = None
                if len(tds) > 6:
                    try:
                        cap_text = tds[6].text.strip().replace(',', '')
                        market_cap = int(cap_text) if cap_text.isdigit() else None
                    except:
                        pass
                
                stocks.append({
                    'code': code, 'name': name, 'price': price,
                    'change_pct': change_pct, 'market_cap': market_cap, 'market': market
                })
            except:
                continue
        
        time.sleep(0.2)
    
    log(f"  네이버 {market}: {len(stocks)}개 로드")
    return stocks

def get_naver_stock_detail(code):
    """네이버에서 종목 상세 정보 - 확장 버전"""
    data = {}

    # 1. 메인 페이지에서 기본 정보
    url = f"https://finance.naver.com/item/main.naver?code={code}"
    resp = fetch_naver(url)

    if resp:
        try:
            soup = BeautifulSoup(resp.text, 'lxml')

            # PER, PBR, 배당수익률
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

            # 현재가
            try:
                price_elem = soup.select_one('p.no_today span.blind')
                if price_elem:
                    data['price'] = int(price_elem.text.replace(',', ''))
            except:
                pass

            # 시가총액
            try:
                for em in soup.select('em#_market_sum'):
                    cap_text = em.text.strip().replace(',', '').replace('조', '').replace('억', '')
                    # 시가총액은 억 단위로 저장
                    if '조' in em.text:
                        data['market_cap'] = int(float(cap_text) * 10000)
                    else:
                        data['market_cap'] = int(cap_text)
            except:
                pass

            # 외국인 비율
            try:
                for table in soup.select('table'):
                    text = table.get_text()
                    if '외국인' in text:
                        match = re.search(r'외국인[^%\d]*([\d.]+)\s*%', text)
                        if match:
                            val = float(match.group(1))
                            if 0 <= val <= 100:
                                data['foreign_ratio'] = val
                                break
            except:
                pass

            # 52주 고저
            try:
                for table in soup.select('table'):
                    for tr in table.select('tr'):
                        th = tr.select_one('th')
                        td = tr.select_one('td')
                        if th and td:
                            th_text = th.get_text().strip()
                            td_text = td.get_text().strip().replace(',', '')
                            if '52주최고' in th_text or '52주 최고' in th_text:
                                match = re.search(r'(\d+)', td_text)
                                if match:
                                    data['high_52w'] = int(match.group(1))
                            elif '52주최저' in th_text or '52주 최저' in th_text:
                                match = re.search(r'(\d+)', td_text)
                                if match:
                                    data['low_52w'] = int(match.group(1))
            except:
                pass

            # 업종(섹터)
            try:
                for a in soup.select('a[href*="upjong"]'):
                    sector_text = a.get_text().strip()
                    if sector_text and len(sector_text) > 1:
                        data['sector'] = sector_text
                        break
            except:
                pass

        except:
            pass

    # 2. 투자지표 페이지에서 ROE, ROA 등 가져오기 (sleep 제거)
    url2 = f"https://finance.naver.com/item/coinfo.naver?code={code}&target=finsum_more"
    resp2 = fetch_naver(url2)

    if resp2:
        try:
            soup2 = BeautifulSoup(resp2.text, 'lxml')

            # iframe 내용이 있을 수 있으므로 테이블에서 직접 추출
            for table in soup2.select('table'):
                rows = table.select('tr')
                for row in rows:
                    cells = row.select('td, th')
                    if len(cells) >= 2:
                        label = cells[0].get_text().strip()

                        # 가장 최근 값 (보통 마지막 td)
                        for cell in reversed(cells[1:]):
                            val_text = cell.get_text().strip().replace(',', '').replace('%', '')
                            if val_text and val_text != '-' and val_text != 'N/A':
                                try:
                                    val = float(val_text)
                                    if 'ROE' in label and 'roe' not in data:
                                        if -100 < val < 200:
                                            data['roe'] = val
                                    elif 'ROA' in label and 'roa' not in data:
                                        if -100 < val < 100:
                                            data['roa'] = val
                                    elif '영업이익률' in label and 'op_margin' not in data:
                                        if -100 < val < 100:
                                            data['op_margin'] = val
                                    elif '순이익률' in label and 'net_margin' not in data:
                                        if -100 < val < 100:
                                            data['net_margin'] = val
                                    elif '부채비율' in label and 'debt_ratio' not in data:
                                        if 0 <= val < 1000:
                                            data['debt_ratio'] = val
                                    elif '유동비율' in label and 'current_ratio' not in data:
                                        if 0 < val < 1000:
                                            data['current_ratio'] = val / 100  # 백분율 → 배수
                                    elif '매출액증가율' in label and 'revenue_growth' not in data:
                                        if -100 < val < 500:
                                            data['revenue_growth'] = val
                                    elif '영업이익증가율' in label and 'op_growth' not in data:
                                        if -200 < val < 1000:
                                            data['op_growth'] = val
                                    # ★ v3.2.0 추가
                                    elif 'ROIC' in label and 'roic' not in data:
                                        if -100 < val < 200:
                                            data['roic'] = val
                                    elif '이자보상' in label and 'interest_coverage' not in data:
                                        if val > -100:
                                            data['interest_coverage'] = val
                                    elif 'EPS' in label and '증가' in label and 'eps_growth' not in data:
                                        if -500 < val < 1000:
                                            data['eps_growth'] = val
                                    break
                                except:
                                    pass
        except:
            pass

    return data

# ============================================================
# FnGuide 스크래핑
# ============================================================
def get_fnguide_data(code):
    """FnGuide에서 재무 데이터 가져오기"""
    global FNGUIDE_AVAILABLE
    
    # 이미 차단된 경우 스킵
    if FNGUIDE_AVAILABLE is False:
        return {}
    
    data = {}
    url = f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A{code}"

    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT_SHORT)
        if resp.status_code == 403:
            FNGUIDE_AVAILABLE = False
            log("  ⚠️ FnGuide 접근 차단됨")
            return data
        if resp.status_code != 200:
            return data
        
        FNGUIDE_AVAILABLE = True
        soup = BeautifulSoup(resp.text, 'lxml')

        # 시가총액, PER, PBR 등 기본 정보
        try:
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
                                        elif 'ROA' in label and 'roa' not in data:
                                            if -100 < val < 100:
                                                data['roa'] = val
                                    except:
                                        pass
        except:
            pass

        # 재무비율 테이블에서 추가 정보
        try:
            for table in soup.select('table'):
                text = table.get_text()
                if '영업이익률' in text or '부채비율' in text:
                    for tr in table.select('tr'):
                        cells = tr.select('td, th')
                        if len(cells) >= 2:
                            label = cells[0].get_text().strip()

                            # 최신 값 추출 (마지막 유효한 값)
                            for cell in reversed(cells[1:]):
                                val_text = cell.get_text().strip().replace(',', '').replace('%', '')
                                if val_text and val_text != '-' and val_text != 'N/A':
                                    try:
                                        val = float(val_text)
                                        if '영업이익률' in label and 'op_margin' not in data:
                                            data['op_margin'] = val
                                        elif '순이익률' in label and 'net_margin' not in data:
                                            data['net_margin'] = val
                                        elif '부채비율' in label and 'debt_ratio' not in data:
                                            if 0 <= val < 1000:
                                                data['debt_ratio'] = val
                                        elif '유동비율' in label and 'current_ratio' not in data:
                                            if val > 0:
                                                data['current_ratio'] = val / 100
                                        elif '매출액증가율' in label and 'revenue_growth' not in data:
                                            data['revenue_growth'] = val
                                        elif 'EPS' in label and 'eps' not in data:
                                            data['eps'] = val
                                        elif 'BPS' in label and 'bps' not in data:
                                            data['bps'] = val
                                        break
                                    except:
                                        pass
        except:
            pass

    except:
        pass

    return data

def get_naver_etf_list(max_pages=5):
    """네이버에서 ETF 리스트"""
    etfs = []
    log("  네이버 ETF 리스트 로드 시도...")
    
    for page in range(1, max_pages + 1):
        url = f"https://finance.naver.com/sise/etf.naver?page={page}"
        resp = fetch_naver(url)
        
        if not resp:
            if NAVER_AVAILABLE is False:
                return []
            continue
        
        soup = BeautifulSoup(resp.text, 'lxml')
        
        for table in soup.select('table.type_1, table.type_2'):
            for tr in table.select('tr'):
                tds = tr.select('td')
                if len(tds) < 5:
                    continue
                
                try:
                    link = tds[0].select_one('a')
                    if not link:
                        continue
                    
                    href = link.get('href', '')
                    code_match = re.search(r'code=(\d{6})', href)
                    if not code_match:
                        continue
                    
                    code = code_match.group(1)
                    name = link.text.strip()
                    
                    price = None
                    if len(tds) > 1:
                        price_text = tds[1].text.strip().replace(',', '')
                        price = int(price_text) if price_text.isdigit() else None
                    
                    etfs.append({'code': code, 'name': name, 'price': price})
                except:
                    continue
        
        time.sleep(0.1)
    
    log(f"  네이버 ETF: {len(etfs)}개 로드")
    return etfs

def get_krx_etf_data():
    """KRX에서 ETF 전종목 데이터 직접 수집 (가장 정확)"""
    etf_data = {}
    
    try:
        log("  KRX에서 ETF 전종목 로드 중...")
        
        # 1. ETF 전종목 시세
        gen_otp_url = 'http://data.krx.co.kr/comm/fileDn/GenerateOTP/generate.cmd'
        gen_otp_data = {
            'locale': 'ko_KR',
            'mktId': 'ETF',
            'trdDd': datetime.now().strftime('%Y%m%d'),
            'share': '1',
            'money': '1',
            'csvxls_isNo': 'false',
            'name': 'fileDown',
            'url': 'dbms/MDC/STAT/standard/MDCSTAT04301'
        }
        
        otp_resp = requests.post(gen_otp_url, data=gen_otp_data, headers=HEADERS, timeout=10)
        otp = otp_resp.text
        
        down_url = 'http://data.krx.co.kr/comm/fileDn/download_csv/download.cmd'
        down_resp = requests.post(
            down_url, 
            data={'code': otp}, 
            headers={**HEADERS, 'Referer': gen_otp_url},
            timeout=30
        )
        
        from io import StringIO
        csv_text = down_resp.content.decode('euc-kr', errors='ignore')
        df = pd.read_csv(StringIO(csv_text))
        
        if not df.empty and len(df) > 50:
            log(f"  KRX: {len(df)}개 ETF 로드")
            
            for _, row in df.iterrows():
                code = str(row.get('종목코드', '')).strip()
                if not code or len(code) != 6:
                    continue
                    
                etf_data[code] = {
                    'name': str(row.get('종목명', '')).strip(),
                    'price': row.get('종가', row.get('현재가', None)),
                    'nav': row.get('NAV', row.get('순자산가치', None)),
                    'volume': row.get('거래량', None),
                }
            
            return etf_data
                
    except Exception as e:
        log(f"  ⚠️ KRX 크롤링 실패: {e}")
    
    return {}

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

# ★★★ v3.0 수정: SharpeRatio 계산 조건 완화 ★★★
def calc_sharpe_ratio(prices, min_period=200, risk_free=0.04):
    """
    SharpeRatio 계산 - v3.0 수정
    
    변경사항:
    - 기존: period=252 (252일 미만이면 None)
    - 수정: min_period=200 (200일 이상이면 계산)
    - v3.0.1: risk_free 0.02 → 0.04 (4%)로 통일
    
    이유: yfinance period="1y"가 실제로 약 250-251일 반환
    """
    try:
        returns = prices.pct_change().dropna()
        if len(returns) < min_period:  # ★ 252 → 200으로 완화
            return None
        mean_return = returns.mean() * 252
        std_return = returns.std() * np.sqrt(252)
        if std_return == 0:
            return None
        return (mean_return - risk_free) / std_return
    except:
        return None

# ★★★ v3.0 추가: 기존 데이터 기반 SharpeRatio 계산 ★★★
def calc_sharpe_from_existing(return_pct, volatility_pct, period_days=120, risk_free_rate=4.0):
    """
    이미 수집된 수익률/변동성으로 SharpeRatio 계산
    
    Parameters:
    - return_pct: Return120D(%) 또는 Return6M(%)
    - volatility_pct: Volatility20D (일간 변동성 연환산 %)
    - risk_free_rate: 무위험 수익률 (연 4% 가정)
    """
    try:
        if return_pct is None or volatility_pct is None:
            return None
        if pd.isna(return_pct) or pd.isna(volatility_pct):
            return None
        if volatility_pct == 0:
            return None
        
        # 연환산 수익률
        annual_return = float(return_pct) * (252 / period_days)
        
        # 변동성은 이미 연환산된 값 (Volatility20D)
        annual_vol = float(volatility_pct)
        
        # Sharpe Ratio
        sharpe = (annual_return - risk_free_rate) / annual_vol
        return fmt(sharpe, 3)
    except:
        return None

# ============================================================
# 데이터 검증
# ============================================================
def validate_foreign_ratio(value):
    if value is None:
        return None
    try:
        v = float(value)
        if v < 0 or v > 100:
            return None
        if v > 70:
            return None
        return v
    except:
        return None

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
# ★★★ v3.0 추가: 기술적 지표 공통 계산 함수 ★★★
# ============================================================
def add_technical_indicators(row, close, include_ma60_120=False):
    """
    기술적 지표를 row에 추가하는 공통 함수
    
    Parameters:
    - row: 데이터 딕셔너리
    - close: pandas Series (종가)
    - include_ma60_120: MA60/MA120 포함 여부 (US_Stocks용)
    """
    if close is None or len(close) < 20:
        return row
    
    price = close.iloc[-1]
    
    # 수익률 계산 - ★ v3.0: 조건 완화 (252 → 245)
    if len(close) >= 2:
        row['Return1D(%)'] = fmt((close.iloc[-1] / close.iloc[-2] - 1) * 100)
    if len(close) >= 6:
        row['Return1W(%)'] = fmt((close.iloc[-1] / close.iloc[-6] - 1) * 100)
        row['Return5D(%)'] = row['Return1W(%)']  # ★ v3.2.0 추가
    if len(close) >= 22:
        row['Return1M(%)'] = fmt((close.iloc[-1] / close.iloc[-22] - 1) * 100)
        row['Return20D(%)'] = row['Return1M(%)']  # ★ v3.2.0 추가
    if len(close) >= 66:
        row['Return3M(%)'] = fmt((close.iloc[-1] / close.iloc[-66] - 1) * 100)
        row['Return60D(%)'] = row['Return3M(%)']  # PWA 호환
    if len(close) >= 132:
        row['Return6M(%)'] = fmt((close.iloc[-1] / close.iloc[-132] - 1) * 100)
        row['Return120D(%)'] = row['Return6M(%)']  # PWA 호환
    
    # ★ v3.0 수정: Return1Y 조건 완화 (252 → 245)
    # ★ v3.0.1 버그 수정: year_ago_idx 계산 (len-1 → len)
    if len(close) >= 245:
        year_ago_idx = min(len(close), 252)  # ★ 수정: -1 제거
        row['Return1Y(%)'] = fmt((close.iloc[-1] / close.iloc[-year_ago_idx] - 1) * 100)
        row['Return250D(%)'] = row['Return1Y(%)']  # PWA 호환
    
    # 이동평균
    row['MA20'] = fmt(close.rolling(20).mean().iloc[-1])
    row['MA50'] = fmt(close.rolling(50).mean().iloc[-1])
    row['MA200'] = fmt(close.rolling(200).mean().iloc[-1]) if len(close) >= 200 else None
    
    # ★ v3.0 추가: MA60, MA120 (US_Stocks용)
    if include_ma60_120:
        if len(close) >= 60:
            row['MA60'] = fmt(close.rolling(60).mean().iloc[-1])
        if len(close) >= 120:
            row['MA120'] = fmt(close.rolling(120).mean().iloc[-1])
    
    # vs_MA 계산
    if row.get('MA20'): row['vs_MA20(%)'] = fmt((price / row['MA20'] - 1) * 100)
    if row.get('MA50'): row['vs_MA50(%)'] = fmt((price / row['MA50'] - 1) * 100)
    if row.get('MA200'): row['vs_MA200(%)'] = fmt((price / row['MA200'] - 1) * 100)
    
    # ★ v3.0 추가: vs_MA60, vs_MA120
    if include_ma60_120:
        if row.get('MA60'): row['vs_MA60(%)'] = fmt((price / row['MA60'] - 1) * 100)
        if row.get('MA120'): row['vs_MA120(%)'] = fmt((price / row['MA120'] - 1) * 100)
    
    # RSI, BB
    row['RSI14'] = fmt(calc_rsi(close, 14))
    row['BB_Position'] = fmt(calc_bollinger_position(close, 20))
    
    # ★ v3.2.0 추가: MACD
    if len(close) >= 26:
        ema12 = close.ewm(span=12, adjust=False).mean()
        ema26 = close.ewm(span=26, adjust=False).mean()
        macd_line = ema12 - ema26
        signal_line = macd_line.ewm(span=9, adjust=False).mean()
        row['MACD'] = fmt(macd_line.iloc[-1])
        row['MACD_Signal'] = fmt(signal_line.iloc[-1])
        row['MACD_Hist'] = fmt(macd_line.iloc[-1] - signal_line.iloc[-1])
    
    # ★ v3.2.0 추가: ADX (Average Directional Index)
    if hist is not None and len(hist) >= 14:
        try:
            high = hist['High'] if 'High' in hist.columns else None
            low = hist['Low'] if 'Low' in hist.columns else None
            if high is not None and low is not None:
                # True Range
                tr1 = high - low
                tr2 = abs(high - close.shift(1))
                tr3 = abs(low - close.shift(1))
                tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
                atr = tr.rolling(14).mean()
                
                # +DM, -DM
                up_move = high - high.shift(1)
                down_move = low.shift(1) - low
                plus_dm = np.where((up_move > down_move) & (up_move > 0), up_move, 0)
                minus_dm = np.where((down_move > up_move) & (down_move > 0), down_move, 0)
                
                plus_di = 100 * pd.Series(plus_dm).rolling(14).mean() / atr
                minus_di = 100 * pd.Series(minus_dm).rolling(14).mean() / atr
                
                dx = 100 * abs(plus_di - minus_di) / (plus_di + minus_di)
                adx = dx.rolling(14).mean()
                
                row['ADX'] = fmt(adx.iloc[-1])
        except:
            pass
    
    # 52주 고저
    if len(close) >= 245:  # ★ v3.0: 252 → 245
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
    
    # ★ v3.0 수정: SharpeRatio (조건 완화)
    row['SharpeRatio'] = fmt(calc_sharpe_ratio(close))
    
    # ★ v3.0 추가: SharpeRatio 폴백 (가격 데이터 부족 시 기존 데이터로 계산)
    if row.get('SharpeRatio') is None:
        row['SharpeRatio'] = calc_sharpe_from_existing(
            row.get('Return120D(%)'), 
            row.get('Volatility20D')
        )
    
    return row

# ============================================================
# ★ v3.0.2 추가: FinanceDataReader 데이터 수집 함수
# ============================================================
def fetch_fdr_stock_data(ticker, start_date=None):
    """FinanceDataReader에서 주식 데이터 가져오기"""
    if not FDR_AVAILABLE:
        return None, pd.DataFrame()
    
    try:
        if start_date is None:
            start_date = (datetime.now() - timedelta(days=400)).strftime('%Y-%m-%d')
        
        df = fdr.DataReader(ticker, start_date)
        if df is not None and not df.empty:
            return df, True
    except:
        pass
    return None, False

def fetch_fdr_stock_list(market='KOSPI'):
    """FinanceDataReader에서 종목 리스트 가져오기"""
    if not FDR_AVAILABLE:
        return []
    
    try:
        stocks = fdr.StockListing(market)
        if stocks is not None and not stocks.empty:
            result = []
            for _, row in stocks.iterrows():
                code = str(row.get('Code', row.get('Symbol', ''))).zfill(6)
                name = row.get('Name', row.get('종목명', ''))
                if code and name:
                    result.append((code, name, market))
            return result
    except:
        pass
    return []

def fetch_fdr_etf_list():
    """FinanceDataReader에서 ETF 리스트 가져오기"""
    if not FDR_AVAILABLE:
        return []
    
    try:
        etfs = fdr.StockListing('ETF/KR')
        if etfs is not None and not etfs.empty:
            result = []
            for _, row in etfs.iterrows():
                code = str(row.get('Code', row.get('Symbol', ''))).zfill(6)
                name = row.get('Name', row.get('종목명', ''))
                if code and name:
                    result.append(code)
            return result
    except:
        pass
    return []

# ============================================================
# 한국 주식 수집
# ============================================================

def collect_kr_stock_codes():
    """
    ★ v3.2.0: 모든 소스에서 종목 코드 수집 후 통합
    Returns: dict {code: {'name': str, 'market': str}}
    """
    all_stocks = {}  # {code: {'name': name, 'market': market}}
    
    log("  [1단계] 종목 리스트 통합 수집...")
    
    # ----------------------------------------
    # 소스 1: pykrx (가장 안정적)
    # ----------------------------------------
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
    
    # ----------------------------------------
    # 소스 2: KRX 직접 크롤링
    # ----------------------------------------
    try:
        log("    KRX 직접 로드 중...")
        krx_url = "http://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"
        krx_headers = {
            'User-Agent': 'Mozilla/5.0',
            'Content-Type': 'application/x-www-form-urlencoded',
        }
        krx_data = {
            'bld': 'dbms/MDC/STAT/standard/MDCSTAT01501',
            'mktId': 'STK',
            'share': '1',
        }
        resp = requests.post(krx_url, data=krx_data, headers=krx_headers, timeout=10)
        if resp.status_code == 200:
            result = resp.json()
            items = result.get('OutBlock_1', [])
            added = 0
            for item in items:
                code = str(item.get('ISU_SRT_CD', '')).zfill(6)
                name = item.get('ISU_ABBRV', '')
                if len(code) == 6 and code.isdigit():
                    if code not in all_stocks:
                        all_stocks[code] = {'name': name, 'market': 'KOSPI'}
                        added += 1
                    elif not all_stocks[code].get('name'):
                        all_stocks[code]['name'] = name
            log(f"    KRX: {added}개 추가 → 총 {len(all_stocks)}개")
    except Exception as e:
        log(f"    ⚠️ KRX 실패: {str(e)[:30]}")
    
    # ----------------------------------------
    # 소스 3: FinanceDataReader
    # ----------------------------------------
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
    
    # ----------------------------------------
    # 소스 4: 네이버 금융
    # ----------------------------------------
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
                    elif not all_stocks[code].get('name'):
                        all_stocks[code]['name'] = name
            log(f"    네이버: {added}개 추가 → 총 {len(all_stocks)}개")
        except Exception as e:
            log(f"    ⚠️ 네이버 실패: {str(e)[:30]}")
    
    # ----------------------------------------
    # 소스 5: NXT
    # ----------------------------------------
    if NXT_AVAILABLE is not False:
        try:
            log("    NXT 로드 중...")
            nxt_data = get_nxt_cached_data()
            added = 0
            for code, info in nxt_data.items():
                code = str(code).zfill(6)
                name = info.get('name', '')
                if len(code) == 6 and code.isdigit():
                    if code not in all_stocks:
                        all_stocks[code] = {'name': name, 'market': info.get('market', 'KOSPI')}
                        added += 1
                    elif not all_stocks[code].get('name'):
                        all_stocks[code]['name'] = name
            log(f"    NXT: {added}개 추가 → 총 {len(all_stocks)}개")
        except Exception as e:
            log(f"    ⚠️ NXT 실패: {str(e)[:30]}")
    
    # ----------------------------------------
    # 소스 6: 하드코딩 (최종 보충)
    # ----------------------------------------
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
    """한국 주식 데이터 수집 - v3.2.0: 종목 리스트 통합 + 시세 순차 수집"""
    log("\n[1/5] 한국 주식 수집 중...")
    
    # ========================================
    # 1단계: 종목 리스트 통합 수집
    # ========================================
    all_stocks = collect_kr_stock_codes()
    
    if not all_stocks:
        log("  ❌ 종목 리스트를 가져올 수 없음")
        return pd.DataFrame()
    
    # 상위 N개 선택 (시총순 정렬이 어려우므로 일단 그대로)
    stock_list = list(all_stocks.items())
    if TOP_N_KR and len(stock_list) > TOP_N_KR:
        stock_list = stock_list[:TOP_N_KR]
    
    log(f"  대상: {len(stock_list)}개")
    
    # ========================================
    # 2단계: 시세 데이터 일괄 로드 (효율성)
    # ========================================
    log("  [2단계] 시세 데이터 일괄 로드...")
    
    # 2-1. NXT 캐시
    nxt_data = get_nxt_cached_data()
    if nxt_data:
        log(f"    NXT: {len(nxt_data)}개")
    
    # 2-2. KRX 전종목 시세 (한번에 가져오기)
    krx_data = {}
    try:
        krx_url = "http://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"
        krx_headers = {
            'User-Agent': 'Mozilla/5.0',
            'Content-Type': 'application/x-www-form-urlencoded',
        }
        krx_params = {
            'bld': 'dbms/MDC/STAT/standard/MDCSTAT01501',
            'mktId': 'STK',
            'share': '1',
        }
        resp = requests.post(krx_url, data=krx_params, headers=krx_headers, timeout=10)
        if resp.status_code == 200:
            result = resp.json()
            for item in result.get('OutBlock_1', []):
                code = str(item.get('ISU_SRT_CD', '')).zfill(6)
                if len(code) == 6 and code.isdigit():
                    krx_data[code] = {
                        'price': int(item.get('TDD_CLSPRC', '0').replace(',', '') or 0),
                        'volume': int(item.get('ACC_TRDVOL', '0').replace(',', '') or 0),
                        'market_cap': int(item.get('MKTCAP', '0').replace(',', '') or 0) // 100000000,  # 억원
                        'change_rate': float(item.get('FLUC_RT', '0').replace(',', '') or 0),
                        'high': int(item.get('TDD_HGPRC', '0').replace(',', '') or 0),
                        'low': int(item.get('TDD_LWPRC', '0').replace(',', '') or 0),
                        'per': float(item.get('PER', '0').replace(',', '') or 0) if item.get('PER') else None,
                        'pbr': float(item.get('PBR', '0').replace(',', '') or 0) if item.get('PBR') else None,
                    }
            log(f"    KRX: {len(krx_data)}개")
    except Exception as e:
        log(f"    ⚠️ KRX 시세 실패: {str(e)[:30]}")
    
    # 2-3. pykrx 전종목 시세
    pykrx_data = {}
    if PYKRX_AVAILABLE:
        try:
            today = datetime.now().strftime("%Y%m%d")
            df = pykrx_stock.get_market_ohlcv(today, market="KOSPI")
            if df is not None and not df.empty:
                for ticker, row_data in df.iterrows():
                    code = str(ticker).zfill(6)
                    pykrx_data[code] = {
                        'price': int(row_data.get('종가', 0)),
                        'volume': int(row_data.get('거래량', 0)),
                        'high': int(row_data.get('고가', 0)),
                        'low': int(row_data.get('저가', 0)),
                        'open': int(row_data.get('시가', 0)),
                        'change_rate': float(row_data.get('등락률', 0)),
                    }
                log(f"    pykrx: {len(pykrx_data)}개")
        except Exception as e:
            log(f"    ⚠️ pykrx 시세 실패: {str(e)[:30]}")
    
    # 2-4. FDR 시세 (히스토리 기반)
    fdr_data = {}
    if FDR_AVAILABLE:
        try:
            # FDR은 개별 호출이 필요하므로 여기서는 스킵
            # 개별 종목 순회 시 필요하면 호출
            pass
        except:
            pass
    
    # ========================================
    # 3단계: 개별 종목 데이터 수집
    # ========================================
    log("  [3단계] 개별 종목 데이터 병합...")
    results = []
    start_time = time.time()
    
    for i, (ticker, info) in enumerate(stock_list):
        try:
            name = info.get('name', '')
            market = info.get('market', 'KOSPI')
            row = {'Code': ticker, 'Name': name, 'Market': market}

            # ========================================
            # 시세 1순위: NXT
            # ========================================
            nxt_info = nxt_data.get(ticker, {})
            if nxt_info:
                if nxt_info.get('price'):
                    row['Price'] = nxt_info['price']
                if nxt_info.get('volume'):
                    row['Volume'] = nxt_info['volume']
                if nxt_info.get('change_rate'):
                    row['Return1D(%)'] = fmt(nxt_info['change_rate'])
                if nxt_info.get('high'):
                    row['DayHigh'] = nxt_info['high']
                if nxt_info.get('low'):
                    row['DayLow'] = nxt_info['low']
                row['DataSource'] = 'NXT'

            # ========================================
            # 시세 2순위: KRX (캐시에서)
            # ========================================
            krx_info = krx_data.get(ticker, {})
            if krx_info:
                if not row.get('Price') and krx_info.get('price'):
                    row['Price'] = krx_info['price']
                if not row.get('Volume') and krx_info.get('volume'):
                    row['Volume'] = krx_info['volume']
                if not row.get('Return1D(%)') and krx_info.get('change_rate'):
                    row['Return1D(%)'] = fmt(krx_info['change_rate'])
                if not row.get('MarketCap(억)') and krx_info.get('market_cap'):
                    row['MarketCap(억)'] = krx_info['market_cap']
                if not row.get('PER') and krx_info.get('per'):
                    row['PER'] = fmt(krx_info['per'])
                if not row.get('PBR') and krx_info.get('pbr'):
                    row['PBR'] = fmt(krx_info['pbr'])
                if not row.get('DataSource'):
                    row['DataSource'] = 'KRX'

            # ========================================
            # 시세 3순위: pykrx (캐시에서)
            # ========================================
            pykrx_info = pykrx_data.get(ticker, {})
            if pykrx_info:
                if not row.get('Price') and pykrx_info.get('price'):
                    row['Price'] = pykrx_info['price']
                if not row.get('Volume') and pykrx_info.get('volume'):
                    row['Volume'] = pykrx_info['volume']
                if not row.get('Return1D(%)') and pykrx_info.get('change_rate'):
                    row['Return1D(%)'] = fmt(pykrx_info['change_rate'])
                if not row.get('DataSource'):
                    row['DataSource'] = 'pykrx'

            # ========================================
            # 시세 4순위: 네이버 (개별 호출 - 재무지표 포함)
            # ========================================
            naver_data = {}
            if NAVER_AVAILABLE is not False:
                naver_data = get_naver_stock_detail(ticker)
                if naver_data:
                    if not row.get('Price') and naver_data.get('price'):
                        row['Price'] = naver_data['price']
                    if not row.get('MarketCap(억)') and naver_data.get('market_cap'):
                        row['MarketCap(억)'] = naver_data['market_cap']
                    if not row.get('Sector') and naver_data.get('sector'):
                        row['Sector'] = naver_data['sector']
                    if not row.get('PER') and naver_data.get('per'):
                        row['PER'] = fmt(naver_data['per'])
                    if not row.get('PBR') and naver_data.get('pbr'):
                        row['PBR'] = fmt(naver_data['pbr'])
                    if not row.get('EPS') and naver_data.get('eps'):
                        row['EPS'] = fmt(naver_data['eps'], 0)
                    if not row.get('BPS') and naver_data.get('bps'):
                        row['BPS'] = fmt(naver_data['bps'], 0)
                    if not row.get('ROE(%)') and naver_data.get('roe'):
                        row['ROE(%)'] = fmt(naver_data['roe'])
                    if not row.get('ROA(%)') and naver_data.get('roa'):
                        row['ROA(%)'] = fmt(naver_data['roa'])
                    if not row.get('OpMargin(%)') and naver_data.get('op_margin'):
                        row['OpMargin(%)'] = fmt(naver_data['op_margin'])
                    if not row.get('NetMargin(%)') and naver_data.get('net_margin'):
                        row['NetMargin(%)'] = fmt(naver_data['net_margin'])
                    if not row.get('RevenueGrowth(%)') and naver_data.get('revenue_growth'):
                        row['RevenueGrowth(%)'] = fmt(naver_data['revenue_growth'])
                    if not row.get('EarningsGrowth(%)') and naver_data.get('op_growth'):
                        row['EarningsGrowth(%)'] = fmt(naver_data['op_growth'])
                    if not row.get('DebtRatio(%)') and naver_data.get('debt_ratio'):
                        row['DebtRatio(%)'] = fmt(naver_data['debt_ratio'])
                    if not row.get('CurrentRatio') and naver_data.get('current_ratio'):
                        row['CurrentRatio'] = fmt(naver_data['current_ratio'])
                    if not row.get('ForeignRatio(%)') and naver_data.get('foreign_ratio'):
                        row['ForeignRatio(%)'] = fmt(naver_data['foreign_ratio'])
                    if not row.get('DivYield(%)') and naver_data.get('div_yield'):
                        row['DivYield(%)'] = fmt(naver_data['div_yield'])
                    if not row.get('52wHigh') and naver_data.get('high_52w'):
                        row['52wHigh'] = naver_data['high_52w']
                    if not row.get('52wLow') and naver_data.get('low_52w'):
                        row['52wLow'] = naver_data['low_52w']
                    # ★ v3.2.0 추가 필드
                    if not row.get('ROIC(%)') and naver_data.get('roic'):
                        row['ROIC(%)'] = fmt(naver_data['roic'])
                    if not row.get('InterestCoverage') and naver_data.get('interest_coverage'):
                        row['InterestCoverage'] = fmt(naver_data['interest_coverage'])
                    if not row.get('EPSGrowth(%)') and naver_data.get('eps_growth'):
                        row['EPSGrowth(%)'] = fmt(naver_data['eps_growth'])
                    if not row.get('DataSource'):
                        row['DataSource'] = 'Naver'

            # ========================================
            # 2. FnGuide (결측값 폴백)
            # ========================================
            need_fnguide = (
                row.get('ROE(%)') is None or
                row.get('ROA(%)') is None or
                row.get('OpMargin(%)') is None or
                row.get('DebtRatio(%)') is None
            )

            if need_fnguide:
                fnguide_data = get_fnguide_data(ticker)
                if fnguide_data:
                    if row.get('PER') is None and fnguide_data.get('per'):
                        row['PER'] = fmt(fnguide_data['per'])
                    if row.get('PBR') is None and fnguide_data.get('pbr'):
                        row['PBR'] = fmt(fnguide_data['pbr'])
                    if row.get('ROE(%)') is None and fnguide_data.get('roe'):
                        row['ROE(%)'] = fmt(fnguide_data['roe'])
                    if row.get('ROA(%)') is None and fnguide_data.get('roa'):
                        row['ROA(%)'] = fmt(fnguide_data['roa'])
                    if row.get('OpMargin(%)') is None and fnguide_data.get('op_margin'):
                        row['OpMargin(%)'] = fmt(fnguide_data['op_margin'])
                    if row.get('NetMargin(%)') is None and fnguide_data.get('net_margin'):
                        row['NetMargin(%)'] = fmt(fnguide_data['net_margin'])
                    if row.get('DebtRatio(%)') is None and fnguide_data.get('debt_ratio'):
                        row['DebtRatio(%)'] = fmt(fnguide_data['debt_ratio'])
                    if row.get('CurrentRatio') is None and fnguide_data.get('current_ratio'):
                        row['CurrentRatio'] = fmt(fnguide_data['current_ratio'])
                    if row.get('RevenueGrowth(%)') is None and fnguide_data.get('revenue_growth'):
                        row['RevenueGrowth(%)'] = fmt(fnguide_data['revenue_growth'])
                    if row.get('EPS') is None and fnguide_data.get('eps'):
                        row['EPS'] = fmt(fnguide_data['eps'], 0)
                    if row.get('BPS') is None and fnguide_data.get('bps'):
                        row['BPS'] = fmt(fnguide_data['bps'], 0)

            # ========================================
            # 3. pykrx (결측값 폴백)
            # ========================================
            if PYKRX_AVAILABLE:
                try:
                    today_str = datetime.now().strftime("%Y%m%d")
                    ohlcv = pykrx_stock.get_market_ohlcv_by_date(
                        (datetime.now() - timedelta(days=7)).strftime("%Y%m%d"),
                        today_str, ticker
                    )
                    if not ohlcv.empty:
                        if row.get('Price') is None:
                            row['Price'] = fmt(ohlcv['종가'].iloc[-1], 0)
                        if row.get('Volume') is None:
                            row['Volume'] = int(ohlcv['거래량'].iloc[-1]) if ohlcv['거래량'].iloc[-1] else None
                except:
                    pass

            # ========================================
            # 4. ★ v3.0.2: FinanceDataReader 우선 (기술적 지표)
            # ========================================
            hist = pd.DataFrame()
            fdr_data = None
            
            if FDR_AVAILABLE:
                try:
                    fdr_data = fdr.DataReader(ticker, DATE_1Y_AGO)
                    if fdr_data is not None and not fdr_data.empty and len(fdr_data) > 20:
                        hist = fdr_data
                        if 'Close' not in hist.columns and '종가' in hist.columns:
                            hist = hist.rename(columns={'종가': 'Close', '시가': 'Open', '고가': 'High', '저가': 'Low', '거래량': 'Volume'})
                        log(f"    {ticker}: FDR 데이터 {len(hist)}일") if i == 0 else None
                except:
                    pass

            # ========================================
            # 5. yfinance (폴백)
            # ========================================
            info = {}
            
            if hist.empty:
                ticker_variants = [f"{ticker}.KS"]
                t, hist, info = try_multiple_tickers(ticker_variants, max_retries=1)
            else:
                # FDR 성공 시에도 yfinance info는 가져옴 (재무 데이터용)
                try:
                    t = yf.Ticker(f"{ticker}.KS")
                    info = t.info or {}
                except:
                    info = {}

            if not row.get('Price'):
                row['Price'] = fmt(safe_get(info, 'regularMarketPrice'))
            if not row.get('Sector'):
                row['Sector'] = safe_get(info, 'sector')
            if row.get('MarketCap(억)') is None:
                cap = safe_get(info, 'marketCap')
                if cap:
                    row['MarketCap(억)'] = fmt(cap / 1e8, 0)
            
            if row.get('PER') is None:
                row['PER'] = fmt(safe_get(info, 'trailingPE'))
            if row.get('ForwardPE') is None:
                row['ForwardPE'] = fmt(safe_get(info, 'forwardPE'))
            if row.get('PBR') is None:
                row['PBR'] = fmt(safe_get(info, 'priceToBook'))
            
            if row.get('ROE(%)') is None and safe_get(info, 'returnOnEquity'):
                row['ROE(%)'] = fmt(safe_get(info, 'returnOnEquity') * 100)
            if row.get('ROA(%)') is None and safe_get(info, 'returnOnAssets'):
                row['ROA(%)'] = fmt(safe_get(info, 'returnOnAssets') * 100)
            if row.get('OpMargin(%)') is None and safe_get(info, 'operatingMargins'):
                row['OpMargin(%)'] = fmt(safe_get(info, 'operatingMargins') * 100)
            if row.get('GrossMargin(%)') is None and safe_get(info, 'grossMargins'):
                row['GrossMargin(%)'] = fmt(safe_get(info, 'grossMargins') * 100)
            
            if row.get('RevenueGrowth(%)') is None and safe_get(info, 'revenueGrowth'):
                row['RevenueGrowth(%)'] = fmt(safe_get(info, 'revenueGrowth') * 100)
            if row.get('EarningsGrowth(%)') is None and safe_get(info, 'earningsGrowth'):
                row['EarningsGrowth(%)'] = fmt(safe_get(info, 'earningsGrowth') * 100)

            if row.get('CurrentRatio') is None:
                row['CurrentRatio'] = fmt(safe_get(info, 'currentRatio'))
            if row.get('DebtRatio(%)') is None:
                row['DebtRatio(%)'] = fmt(safe_get(info, 'debtToEquity'))
            
            if row.get('DivYield(%)') is None and safe_get(info, 'dividendYield'):
                row['DivYield(%)'] = fmt(safe_get(info, 'dividendYield') * 100)
            
            # ★ v3.0: 기술적 지표 공통 함수 사용 (KR_Stocks는 MA60/120 포함)
            if not hist.empty and len(hist) > 20:
                close = hist['Close']
                row = add_technical_indicators(row, close, include_ma60_120=True)

            # 최종 폴백
            kr_field_mapping = {
                'PER': (('trailingPE', 'forwardPE'), 1, 2),
                'PBR': (('priceToBook',), 1, 2),
                'ROE(%)': (('returnOnEquity',), 100, 2),
                'ROA(%)': (('returnOnAssets',), 100, 2),
                'DivYield(%)': (('dividendYield',), 100, 2),
                'Beta': (('beta',), 1, 2),
                'MarketCap(억)': (('marketCap',), 1e-8, 0),
            }
            row = fill_missing_from_info(row, info, kr_field_mapping)

            if row.get('52wHigh') is None:
                val = safe_get(info, 'fiftyTwoWeekHigh')
                if val:
                    row['52wHigh'] = fmt(val)
            if row.get('52wLow') is None:
                val = safe_get(info, 'fiftyTwoWeekLow')
                if val:
                    row['52wLow'] = fmt(val)

            row = calc_data_quality_score(row, REQUIRED_COLS_KR_STOCK, 'kr_stock')
            results.append(row)
            
        except Exception as e:
            results.append({'Code': ticker, 'Name': name, 'Market': market, 'Remark': str(e)[:30]})
        
        if (i + 1) % 10 == 0 or i == 0:
            elapsed = time.time() - start_time
            per_stock = elapsed / (i + 1) if i > 0 else 0
            remaining = per_stock * (len(all_tickers) - i - 1)
            log(f"  진행: {i+1}/{len(all_tickers)} ({(i+1)/len(all_tickers)*100:.0f}%) - 남은시간: {remaining/60:.1f}분")
        
        time.sleep(0.02)
    
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
            
            # 밸류에이션
            row['PER'] = fmt(safe_get(info, 'trailingPE'))
            row['ForwardPE'] = fmt(safe_get(info, 'forwardPE'))
            row['PBR'] = fmt(safe_get(info, 'priceToBook'))
            row['PSR'] = fmt(safe_get(info, 'priceToSalesTrailing12Months'))
            row['PEG'] = fmt(safe_get(info, 'pegRatio'))
            row['EV/EBITDA'] = fmt(safe_get(info, 'enterpriseToEbitda'))
            
            # 수익성
            row['ROE(%)'] = fmt(safe_get(info, 'returnOnEquity', default=0) * 100) if safe_get(info, 'returnOnEquity') else None
            row['ROA(%)'] = fmt(safe_get(info, 'returnOnAssets', default=0) * 100) if safe_get(info, 'returnOnAssets') else None
            row['GrossMargin(%)'] = fmt(safe_get(info, 'grossMargins', default=0) * 100) if safe_get(info, 'grossMargins') else None
            row['OpMargin(%)'] = fmt(safe_get(info, 'operatingMargins', default=0) * 100) if safe_get(info, 'operatingMargins') else None
            row['NetMargin(%)'] = fmt(safe_get(info, 'profitMargins', default=0) * 100) if safe_get(info, 'profitMargins') else None
            
            # FCF Yield
            fcf = safe_get(info, 'freeCashflow')
            cap = safe_get(info, 'marketCap')
            if fcf and cap and cap > 0:
                row['FCFYield(%)'] = fmt(fcf / cap * 100)
            
            # 성장성
            row['RevenueGrowth(%)'] = fmt(safe_get(info, 'revenueGrowth', default=0) * 100) if safe_get(info, 'revenueGrowth') else None
            row['EarningsGrowth(%)'] = fmt(safe_get(info, 'earningsGrowth', default=0) * 100) if safe_get(info, 'earningsGrowth') else None
            
            # 안정성
            row['CurrentRatio'] = fmt(safe_get(info, 'currentRatio'))
            row['QuickRatio'] = fmt(safe_get(info, 'quickRatio'))
            row['DebtRatio(%)'] = fmt(safe_get(info, 'debtToEquity'))
            row['Debt/Equity'] = fmt(safe_get(info, 'debtToEquity'))
            
            # 배당
            row['DivYield(%)'] = fmt(safe_get(info, 'dividendYield', default=0) * 100) if safe_get(info, 'dividendYield') else None
            row['PayoutRatio(%)'] = fmt(safe_get(info, 'payoutRatio', default=0) * 100) if safe_get(info, 'payoutRatio') else None
            
            # 베타
            row['Beta'] = fmt(safe_get(info, 'beta'))
            
            # 기관 보유 비율
            inst = safe_get(info, 'heldPercentInstitutions')
            if inst:
                row['InstOwn(%)'] = fmt(inst * 100)
            
            # ★ v3.0 추가: EPS, BPS, Volume
            row['EPS'] = fmt(safe_get(info, 'trailingEps'))
            row['BPS'] = fmt(safe_get(info, 'bookValue'))
            row['Volume'] = safe_get(info, 'regularMarketVolume') or safe_get(info, 'averageVolume')
            
            # ★ v3.0: 기술적 지표 공통 함수 사용 (MA60/120 포함!)
            if not hist.empty and len(hist) > 20:
                close = hist['Close']
                row = add_technical_indicators(row, close, include_ma60_120=True)  # ★ MA60/120 추가
                
                # YTD 수익률
                try:
                    ytd_start = close[close.index >= f"{datetime.now().year}-01-01"]
                    if len(ytd_start) > 1:
                        row['ReturnYTD(%)'] = fmt((close.iloc[-1] / ytd_start.iloc[0] - 1) * 100)
                except:
                    pass

            # 최종 폴백
            us_field_mapping = {
                'PER': (('trailingPE', 'forwardPE'), 1, 2),
                'ForwardPE': (('forwardPE',), 1, 2),
                'PBR': (('priceToBook',), 1, 2),
                'ROE(%)': (('returnOnEquity',), 100, 2),
                'ROA(%)': (('returnOnAssets',), 100, 2),
                'DivYield(%)': (('dividendYield',), 100, 2),
                'Beta': (('beta',), 1, 2),
                'GrossMargin(%)': (('grossMargins',), 100, 2),
                'OpMargin(%)': (('operatingMargins',), 100, 2),
                'NetMargin(%)': (('profitMargins',), 100, 2),
            }
            row = fill_missing_from_info(row, info, us_field_mapping)

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
            results.append({'Ticker': ticker, 'Remark': str(e)[:30]})

        if (i + 1) % 20 == 0:
            log(f"  진행: {i+1}/{len(tickers)}")

        time.sleep(0.15)

    log(f"  ✅ 완료: {len(results)}개")
    return pd.DataFrame(results)

# ============================================================
# ETF 공통 함수
# ============================================================
def get_etf_data(tickers, region=""):
    """ETF 데이터 수집 공통 함수 - v3.0 수정"""
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
            
            # ★ v3.0 수정: ExpenseRatio 하드코딩 우선
            yf_expense = safe_get(info, 'expenseRatio')
            if yf_expense:
                row['ExpenseRatio(%)'] = fmt(yf_expense * 100, 3)
            elif ticker in US_ETF_EXPENSE:  # ★ 하드코딩 폴백
                row['ExpenseRatio(%)'] = US_ETF_EXPENSE[ticker]
            else:
                row['ExpenseRatio(%)'] = None
            
            row['TotalAssets(B)'] = fmt(safe_get(info, 'totalAssets', default=0) / 1e9, 2) if safe_get(info, 'totalAssets') else None
            row['DivYield(%)'] = fmt(safe_get(info, 'yield', default=0) * 100) if safe_get(info, 'yield') else None
            
            # ★ v3.0: 기술적 지표 공통 함수 사용
            if not hist.empty and len(hist) > 20:
                close = hist['Close']
                row = add_technical_indicators(row, close, include_ma60_120=False)

            # 최종 폴백
            etf_field_mapping = {
                'DivYield(%)': (('yield', 'dividendYield'), 100, 2),
                'TotalAssets(B)': (('totalAssets',), 1e-9, 2),
            }
            row = fill_missing_from_info(row, info, etf_field_mapping)

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
            results.append({'Ticker': ticker, 'Remark': str(e)[:30]})

        time.sleep(0.15)

    return pd.DataFrame(results)

# ============================================================
# 한국 ETF
# ============================================================
def get_korea_etfs():
    """한국 ETF 데이터 - v3.0.2: FDR 우선, 폴백 강화"""
    log("\n[3/5] 한국 ETF 수집 중...")
    
    kr_etf_list = []
    krx_data = {}
    
    # ★ 1순위: pykrx (가장 안정적)
    if PYKRX_AVAILABLE and not kr_etf_list:
        try:
            log("  pykrx에서 ETF 리스트 로드 중...")
            kr_etf_list = pykrx_stock.get_etf_ticker_list()
            if kr_etf_list:
                log(f"  pykrx: {len(kr_etf_list)}개 ETF ✅")
        except Exception as e:
            log(f"  ⚠️ pykrx ETF 실패: {e}")
    
    # ★ 2순위: KRX 직접 크롤링
    if not kr_etf_list:
        try:
            krx_data = get_krx_etf_data()
            if krx_data:
                kr_etf_list = list(krx_data.keys())
                log(f"  KRX: {len(kr_etf_list)}개 ETF ✅")
        except Exception as e:
            log(f"  ⚠️ KRX 크롤링 실패: {e}")
    
    # ★ 3순위: FinanceDataReader
    if not kr_etf_list and FDR_AVAILABLE:
        try:
            log("  FinanceDataReader에서 ETF 리스트 로드 중...")
            kr_etf_list = fetch_fdr_etf_list()
            if kr_etf_list:
                log(f"  FinanceDataReader: {len(kr_etf_list)}개 ETF ✅")
        except Exception as e:
            log(f"  ⚠️ FinanceDataReader ETF 실패: {e}")
    
    # ★ 4순위: 네이버
    if not kr_etf_list and NAVER_AVAILABLE is not False:
        try:
            naver_etfs = get_naver_etf_list(max_pages=10)
            if naver_etfs:
                kr_etf_list = [etf['code'] for etf in naver_etfs]
                log(f"  네이버: {len(kr_etf_list)}개 ETF ✅")
        except Exception as e:
            log(f"  ⚠️ 네이버 ETF 실패: {e}")
    
    # ★ 5순위: 하드코딩 폴백 (항상 작동)
    if not kr_etf_list or len(kr_etf_list) < 50:
        log("  ⚠️ ETF 리스트 부족, 하드코딩 데이터로 보충")
        hardcoded_etfs = list(KR_ETF_EXPENSE.keys())
        # 기존 리스트와 합치기 (중복 제거)
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
            code = str(code).zfill(6)  # 6자리로 패딩
            row = {'Code': code, 'Region': 'KR'}

            # 1. KRX 데이터 먼저 적용
            krx_info = krx_data.get(code, {}) if krx_data else {}
            if krx_info:
                row['Name'] = krx_info.get('name', '')
                if krx_info.get('price'):
                    row['Price'] = fmt(krx_info['price'], 0)
                if krx_info.get('nav'):
                    row['NAV'] = fmt(krx_info['nav'], 0)
                if krx_info.get('volume'):
                    row['Volume'] = int(krx_info['volume'])

            # 2. pykrx 보충
            if PYKRX_AVAILABLE and not row.get('Name'):
                try:
                    row['Name'] = pykrx_stock.get_etf_ticker_name(code)
                    today_str = datetime.now().strftime("%Y%m%d")
                    ohlcv = pykrx_stock.get_etf_ohlcv_by_date(
                        (datetime.now() - timedelta(days=7)).strftime("%Y%m%d"),
                        today_str, code
                    )
                    if not ohlcv.empty:
                        if not row.get('Price'):
                            row['Price'] = fmt(ohlcv['종가'].iloc[-1], 0)
                        if not row.get('Volume'):
                            row['Volume'] = int(ohlcv['거래량'].iloc[-1]) if ohlcv['거래량'].iloc[-1] else None
                except:
                    pass

            # ★ v3.0.2: FinanceDataReader 우선 (기술적 지표)
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

            # 3. yfinance (폴백)
            info = {}
            if hist.empty:
                ticker_variants = [f"{code}.KS", f"{code}.KQ"]
                t, hist, info = try_multiple_tickers(ticker_variants, max_retries=1)
            else:
                try:
                    t = yf.Ticker(f"{code}.KS")
                    info = t.info or {}
                except:
                    info = {}

            if not row.get('Name'):
                row['Name'] = safe_get(info, 'shortName', 'longName') or code
            if not row.get('Price'):
                row['Price'] = fmt(safe_get(info, 'regularMarketPrice'))

            # ★ v3.0 수정: Category 하드코딩 우선
            yf_category = safe_get(info, 'category')
            if yf_category:
                row['Category'] = yf_category
            elif code in KR_ETF_CATEGORY:  # ★ 하드코딩 폴백
                row['Category'] = KR_ETF_CATEGORY[code]
            else:
                row['Category'] = None
            
            # ★ v3.0 수정: ExpenseRatio 하드코딩 우선
            yf_expense = safe_get(info, 'expenseRatio')
            if yf_expense:
                row['ExpenseRatio(%)'] = fmt(yf_expense * 100, 3)
            elif code in KR_ETF_EXPENSE:  # ★ 하드코딩 폴백
                row['ExpenseRatio(%)'] = KR_ETF_EXPENSE[code]
            else:
                row['ExpenseRatio(%)'] = None
            
            # ★ v3.0 추가: DivYield 하드코딩
            yf_yield = safe_get(info, 'yield') or safe_get(info, 'dividendYield')
            if yf_yield:
                row['DivYield(%)'] = fmt(yf_yield * 100, 2)
            elif code in KR_ETF_DIVYIELD:  # ★ 하드코딩 폴백
                row['DivYield(%)'] = KR_ETF_DIVYIELD[code]
            else:
                row['DivYield(%)'] = None
            
            # ★ v3.0 추가: TotalAssets (v3.0.1: 단위 수정 1e12 → 1e9)
            total_assets = safe_get(info, 'totalAssets')
            if total_assets:
                row['TotalAssets(B)'] = fmt(total_assets / 1e9, 2)  # Billion 단위
            
            # ★ v3.0: 기술적 지표 공통 함수 사용
            if not hist.empty and len(hist) > 20:
                close = hist['Close']
                row = add_technical_indicators(row, close, include_ma60_120=False)
            else:
                if not info:
                    continue

            # 최종 폴백
            kr_etf_field_mapping = {
                'DivYield(%)': (('yield', 'dividendYield'), 100, 2),
            }
            row = fill_missing_from_info(row, info, kr_etf_field_mapping)

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
            pass
        
        if (i + 1) % 20 == 0 or i == 0:
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
    """CNN Fear & Greed Index 스크래핑"""
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
    """S&P 500 Shiller CAPE Ratio from multpl.com"""
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
    """S&P 500 Forward PE from SPY ETF"""
    try:
        for ticker in ['SPY', 'IVV', 'VOO']:
            t = yf.Ticker(ticker)
            info = t.info
            pe = safe_get(info, 'trailingPE')
            if pe and 10 < pe < 50:
                return fmt(pe, 2)
    except:
        pass
    return None

def get_ism_pmi():
    """ISM 제조업 PMI from tradingeconomics"""
    try:
        url = "https://tradingeconomics.com/united-states/business-confidence"
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT_LONG)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'lxml')
            for elem in soup.select('#aspnetForm td, #aspnetForm span'):
                text = elem.get_text().strip()
                try:
                    val = float(text)
                    if 40 <= val <= 65:
                        return fmt(val, 1)
                except:
                    pass
    except:
        pass
    
    try:
        url = "https://www.investing.com/economic-calendar/ism-manufacturing-pmi-173"
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT_LONG)
        if resp.status_code == 200:
            match = re.search(r'Actual.*?(\d+\.?\d*)', resp.text)
            if match:
                val = float(match.group(1))
                if 40 <= val <= 65:
                    return fmt(val, 1)
    except:
        pass
    
    return None

def get_korea_market_indicators():
    """한국 시장 지표 (pykrx + 네이버 대안)"""
    indicators = {}

    if PYKRX_AVAILABLE:
        try:
            today_str = datetime.now().strftime("%Y%m%d")

            try:
                from_date = (datetime.now() - timedelta(days=7)).strftime("%Y%m%d")
                kospi_fund = pykrx_stock.get_index_fundamental(from_date, today_str, "1001")
                if not kospi_fund.empty:
                    latest = kospi_fund.iloc[-1]
                    if 'PER' in kospi_fund.columns:
                        indicators['kospi_per'] = fmt(latest['PER'], 2)
                    if 'PBR' in kospi_fund.columns:
                        indicators['kospi_pbr'] = fmt(latest['PBR'], 2)
                    if 'DIV' in kospi_fund.columns:
                        indicators['kospi_div'] = fmt(latest['DIV'], 2)
            except Exception as e:
                log(f"  ⚠️ pykrx 펀더멘털 실패: {e}")

            try:
                from_date = (datetime.now() - timedelta(days=30)).strftime("%Y%m%d")
                investor = pykrx_stock.get_market_trading_value_by_date(from_date, today_str, "KOSPI")
                if not investor.empty and '외국인' in investor.columns:
                    recent_5 = investor['외국인'].tail(5).sum()
                    indicators['foreign_net_buy'] = fmt(recent_5 / 1e8, 0)
                    recent_20 = investor['외국인'].tail(20).sum()
                    indicators['foreign_net_buy_20d'] = fmt(recent_20 / 1e8, 0)
                    
                if not investor.empty:
                    if '개인' in investor.columns:
                        indicators['individual_net_buy'] = fmt(investor['개인'].tail(5).sum() / 1e8, 0)
                    if '기관합계' in investor.columns:
                        indicators['institution_net_buy'] = fmt(investor['기관합계'].tail(5).sum() / 1e8, 0)
            except Exception as e:
                log(f"  ⚠️ pykrx 투자자 실패: {e}")

        except Exception as e:
            log(f"  ⚠️ pykrx 전체 실패: {e}")

    if not indicators.get('kospi_per') and NAVER_AVAILABLE is not False:
        try:
            url = "https://finance.naver.com/sise/sise_index.naver?code=KOSPI"
            resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT_SHORT)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, 'lxml')
                for td in soup.select('td'):
                    text = td.get_text().strip()
                    if 'PER' in text or 'per' in text.lower():
                        parent = td.find_parent('tr')
                        if parent:
                            tds = parent.select('td')
                            for t in tds:
                                val_text = t.get_text().strip().replace(',', '')
                                try:
                                    val = float(val_text)
                                    if 5 < val < 50:
                                        indicators['kospi_per'] = fmt(val, 2)
                                        break
                                except:
                                    pass
                log("  ✅ 네이버에서 코스피 지표 로드")
        except Exception as e:
            log(f"  ⚠️ 네이버 코스피 지표 실패: {e}")

    return indicators

def get_market_indicators():
    """글로벌 시장 지표 - 확장 버전"""
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
        '^STOXX50E': ('유로스톡스50', '지수'),
        '^FTSE': ('FTSE 100', '지수'),

        'USDKRW=X': ('USD/KRW', '환율'),
        'EURUSD=X': ('EUR/USD', '환율'),
        'USDJPY=X': ('USD/JPY', '환율'),
        'DX-Y.NYB': ('달러 인덱스', '환율'),
        'USDCNY=X': ('USD/CNY', '환율'),

        'GC=F': ('금 선물', '원자재'),
        'SI=F': ('은 선물', '원자재'),
        'CL=F': ('WTI 원유', '원자재'),
        'BZ=F': ('브렌트유', '원자재'),
        'NG=F': ('천연가스', '원자재'),
        'HG=F': ('구리 선물', '원자재'),

        '^IRX': ('미국채 3개월', '채권'),
        '^FVX': ('미국채 5년', '채권'),
        '^TNX': ('미국채 10년', '채권'),
        '^TYX': ('미국채 30년', '채권'),

        'BTC-USD': ('비트코인', '암호화폐'),
        'ETH-USD': ('이더리움', '암호화폐'),

        'HYG': ('하이일드 채권 ETF', '신용'),
        'LQD': ('투자등급 채권 ETF', '신용'),
        'TLT': ('장기국채 ETF', '채권'),
        'SHY': ('단기국채 ETF', '채권'),

        '^CPCE': ('Put/Call Ratio', '심리'),
    }

    results = []

    log("  Yahoo Finance 지표 수집 중...")
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
            if len(close) >= 245:  # ★ v3.0: 252 → 245
                row['Return1Y(%)'] = fmt((close.iloc[-1] / close.iloc[-min(len(close), 252)] - 1) * 100)  # ★ v3.0.1 버그 수정
                year_data = close.tail(min(len(close), 252))
                row['52wHigh'] = fmt(year_data.max(), decimals)
                row['52wLow'] = fmt(year_data.min(), decimals)
                row['From52wHigh(%)'] = fmt((current_price / year_data.max() - 1) * 100)
                row['From52wLow(%)'] = fmt((current_price / year_data.min() - 1) * 100)

            results.append(row)
        except:
            pass

        time.sleep(0.05)

    log("  계산 지표 수집 중...")

    try:
        tnx = next((r for r in results if r['Ticker'] == '^TNX'), None)
        irx = next((r for r in results if r['Ticker'] == '^IRX'), None)
        if tnx and irx and tnx.get('Price') and irx.get('Price'):
            spread_10y_3m = float(tnx['Price']) - float(irx['Price'])
            results.append({
                'Category': '스프레드',
                'Ticker': '10Y-3M',
                'Name': '10년-3개월 스프레드',
                'Price': fmt(spread_10y_3m, 3),
                'Signal': '역전' if spread_10y_3m < 0 else '정상',
            })
    except:
        pass

    try:
        tnx = next((r for r in results if r['Ticker'] == '^TNX'), None)
        fvx = next((r for r in results if r['Ticker'] == '^FVX'), None)
        if tnx and fvx and tnx.get('Price') and fvx.get('Price'):
            spread_10y_5y = float(tnx['Price']) - float(fvx['Price'])
            results.append({
                'Category': '스프레드',
                'Ticker': '10Y-5Y',
                'Name': '10년-5년 스프레드',
                'Price': fmt(spread_10y_5y, 3),
            })
    except:
        pass

    try:
        hyg = next((r for r in results if r['Ticker'] == 'HYG'), None)
        lqd = next((r for r in results if r['Ticker'] == 'LQD'), None)
        if hyg and lqd and hyg.get('Return1M(%)') and lqd.get('Return1M(%)'):
            hy_spread_proxy = float(lqd['Return1M(%)']) - float(hyg['Return1M(%)'])
            results.append({
                'Category': '스프레드',
                'Ticker': 'HY-IG',
                'Name': '하이일드 스프레드 (proxy)',
                'Price': fmt(hy_spread_proxy, 2),
                'Description': 'LQD-HYG 1개월 수익률 차이',
            })
    except:
        pass

    log("  심리 지표 수집 중...")

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

    ism_pmi = get_ism_pmi()
    if ism_pmi:
        results.append({
            'Category': '경기',
            'Ticker': 'ISM-PMI',
            'Name': 'ISM 제조업 PMI',
            'Price': ism_pmi,
            'Signal': '확장' if float(ism_pmi) > 50 else '수축',
        })

    log("  한국 시장 지표 수집 중...")

    kr_indicators = get_korea_market_indicators()

    if kr_indicators.get('kospi_per'):
        results.append({
            'Category': '밸류에이션',
            'Ticker': 'KOSPI-PER',
            'Name': '코스피 PER',
            'Price': kr_indicators['kospi_per'],
        })

    if kr_indicators.get('kospi_pbr'):
        results.append({
            'Category': '밸류에이션',
            'Ticker': 'KOSPI-PBR',
            'Name': '코스피 PBR',
            'Price': kr_indicators['kospi_pbr'],
        })

    if kr_indicators.get('kospi_div'):
        results.append({
            'Category': '밸류에이션',
            'Ticker': 'KOSPI-DIV',
            'Name': '코스피 배당수익률',
            'Price': kr_indicators['kospi_div'],
        })

    if kr_indicators.get('foreign_net_buy') is not None:
        val = kr_indicators['foreign_net_buy']
        results.append({
            'Category': '수급',
            'Ticker': 'KR-외국인',
            'Name': '외국인 순매수 (5일, 억원)',
            'Price': val,
            'Signal': '매수' if float(val) > 0 else '매도',
        })

    if kr_indicators.get('foreign_net_buy_20d') is not None:
        val = kr_indicators['foreign_net_buy_20d']
        results.append({
            'Category': '수급',
            'Ticker': 'KR-외국인-20D',
            'Name': '외국인 순매수 (20일 누적, 억원)',
            'Price': val,
            'Signal': '매수' if float(val) > 0 else '매도',
        })

    if kr_indicators.get('individual_net_buy') is not None:
        val = kr_indicators['individual_net_buy']
        results.append({
            'Category': '수급',
            'Ticker': 'KR-개인',
            'Name': '개인 순매수 (5일, 억원)',
            'Price': val,
            'Signal': '매수' if float(val) > 0 else '매도',
        })

    if kr_indicators.get('institution_net_buy') is not None:
        val = kr_indicators['institution_net_buy']
        results.append({
            'Category': '수급',
            'Ticker': 'KR-기관',
            'Name': '기관 순매수 (5일, 억원)',
            'Price': val,
            'Signal': '매수' if float(val) > 0 else '매도',
        })

    log("  추가 ETF 지표 수집 중...")

    additional_etfs = {
        'IEF': ('미국채 7-10년 ETF', '채권'),
        'TIP': ('물가연동채 ETF', '채권'),
        'EMB': ('신흥국 채권 ETF', '신용'),
        'GLD': ('금 ETF', '원자재'),
        'USO': ('원유 ETF', '원자재'),
        'VXX': ('VIX 선물 ETF', '변동성'),
        'SVXY': ('VIX 인버스 ETF', '변동성'),
    }

    for ticker, (name, category) in additional_etfs.items():
        try:
            t = yf.Ticker(ticker)
            hist = t.history(period="3mo")

            if hist.empty:
                continue

            close = hist['Close']
            current_price = close.iloc[-1]

            row = {
                'Category': category,
                'Ticker': ticker,
                'Name': name,
                'Price': fmt(current_price, 2),
            }

            if len(close) >= 22:
                row['Return1M(%)'] = fmt((close.iloc[-1] / close.iloc[-22] - 1) * 100)

            results.append(row)
        except:
            pass

        time.sleep(0.05)

    try:
        tnx = next((r for r in results if r['Ticker'] == '^TNX'), None)
        tip = next((r for r in results if r['Ticker'] == 'TIP'), None)
        if tnx and tip:
            tip_return = tip.get('Return1M(%)')
            if tip_return:
                real_rate = float(tnx['Price']) - (float(tip_return) * 12 / 100)
                results.append({
                    'Category': '금리',
                    'Ticker': 'REAL-RATE',
                    'Name': '실질금리 추정',
                    'Price': fmt(real_rate, 3),
                    'Description': '10Y - TIP implied inflation',
                })
    except:
        pass

    log(f"  ✅ 완료: {len(results)}개")
    return pd.DataFrame(results)

# ============================================================
# Excel 저장
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

# ============================================================
# JSON 저장
# ============================================================
def save_to_json(data_dict, filename):
    log("\nJSON 저장 중...")
    
    output = {
        'metadata': {
            'generated': datetime.now().isoformat(),
            'date': TODAY,
            'version': 'v3.2.0-github-pwa-compatible'  # ★ 버전 업데이트
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
    parser = argparse.ArgumentParser(description='글로벌 주식/ETF 스크리너 (GitHub Actions) v3.0.2')
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
    log("글로벌 주식/ETF 스크리닝 - GitHub Actions v3.0.2")
    log("★ v3.0.2: FinanceDataReader 추가, 데이터 소스 우선순위 개선")
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

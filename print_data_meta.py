#!/usr/bin/env python3
"""
data.json 메타 정보 + 수집 성공률 통계 출력
GitHub Actions의 'Run screener' step에서 호출됨

사용법:
    python scripts/print_data_meta.py [data_path]
    
인자:
    data_path: data.json 경로 (기본: ./data/data.json)
"""
import sys
import json
import os


def main():
    data_path = sys.argv[1] if len(sys.argv) > 1 else './data/data.json'
    
    if not os.path.exists(data_path):
        print(f"❌ {data_path} 파일이 존재하지 않습니다")
        sys.exit(1)
    
    try:
        with open(data_path, encoding='utf-8') as f:
            d = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(f"❌ JSON 파싱 실패: {e}")
        sys.exit(1)
    
    # 메타 정보
    metadata = d.get('metadata', {})
    print('📊 Pipeline version:', metadata.get('version', 'N/A'))
    print('   Date:', metadata.get('date', 'N/A'))
    print('   Generated:', metadata.get('generated', 'N/A'))
    
    data = d.get('data', {})
    meta = data.get('_meta', {})
    if meta:
        print('   Generated at:', meta.get('generated_at', 'N/A'))
        print('   Data sources:', ', '.join(meta.get('data_sources', [])))
        print('   DART enabled:', meta.get('dart_enabled', False))
    
    # 시트별 종목 수
    kr = data.get('KR_Stocks', [])
    us = data.get('US_Stocks', [])
    mi = data.get('Market_Indicators', [])
    print(f'   KR stocks: {len(kr)}')
    print(f'   US stocks: {len(us)}')
    print(f'   Market indicators: {len(mi)}')
    
    # 기타 시트
    other_sheets = [k for k in data.keys() 
                    if k not in ('KR_Stocks', 'US_Stocks', 'Market_Indicators', '_meta')]
    for sheet in other_sheets:
        rows = data.get(sheet, [])
        if isinstance(rows, list):
            print(f'   {sheet}: {len(rows)}')
    
    # v4.4: KR DART/pykrx 수집 성공률
    if kr:
        n = len(kr)
        with_dart = sum(1 for s in kr if 'dart' in str(s.get('_source', '')))
        with_pykrx = sum(1 for s in kr if 'pykrx' in str(s.get('_source', '')))
        with_piotroski = sum(1 for s in kr if s.get('PiotroskiFScore') is not None)
        with_altman = sum(1 for s in kr if s.get('AltmanZScore') is not None)
        with_roic = sum(1 for s in kr if s.get('ROIC(%)') is not None)
        with_shorting = sum(1 for s in kr if s.get('ShortingBalance(억)') is not None)
        with_foreign = sum(1 for s in kr if s.get('Foreign5DNet(억)') is not None)
        
        print(f'   KR - DART 적용: {with_dart}/{n} ({100*with_dart/n:.0f}%)')
        print(f'   KR - pykrx 적용: {with_pykrx}/{n} ({100*with_pykrx/n:.0f}%)')
        print(f'   KR - Piotroski 계산됨: {with_piotroski}/{n} ({100*with_piotroski/n:.0f}%)')
        print(f'   KR - Altman 계산됨: {with_altman}/{n} ({100*with_altman/n:.0f}%)')
        print(f'   KR - ROIC 계산됨: {with_roic}/{n} ({100*with_roic/n:.0f}%)')
        print(f'   KR - 공매도 데이터: {with_shorting}/{n} ({100*with_shorting/n:.0f}%)')
        print(f'   KR - 외국인 순매수: {with_foreign}/{n} ({100*with_foreign/n:.0f}%)')


if __name__ == '__main__':
    main()

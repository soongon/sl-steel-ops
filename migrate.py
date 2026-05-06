#!/usr/bin/env python3
"""
SH철강 운영 데이터 마이그레이션 도구

xlsx와 CSV 사이의 양방향 동기화를 담당.
xlsx는 view, CSV가 truth source.

usage:
    python migrate.py extract                      # xlsx → CSV (데이터 추출)
    python migrate.py rebuild                      # 빈 양식 + CSV → xlsx
    python migrate.py sync                         # extract + rebuild
    python migrate.py extract --xlsx path/to.xlsx  # 다른 xlsx 지정
    python migrate.py rebuild --xlsx path/to.xlsx  # 다른 출력 경로
"""

import csv
import sys
from datetime import datetime, date
from pathlib import Path
from openpyxl import load_workbook

# 동일 패키지의 create_spreadsheets에서 build_workbook 가져오기
sys.path.insert(0, str(Path(__file__).parent / 'workbook'))
from create_spreadsheets import build_workbook  # noqa: E402


# ============================================================
# 시트별 메타데이터
# ============================================================
# header_row: 헤더가 있는 행 (데이터는 그 다음 행부터)
# max_col: 데이터 영역의 최대 컬럼 (영수증 시트는 A~H만, J~L은 합계 영역)
# formula_cols: 자동 수식 컬럼 (1-based) — CSV에 저장하지 않고, rebuild 시 건드리지 않음
# date_cols: 날짜 컬럼 (1-based) — CSV ↔ xlsx 변환 시 datetime 처리
SHEETS = {
    '1.매출': {
        'header_row': 1,
        'max_col': 21,
        'formula_cols': [2, 3, 14, 19, 20],   # 주문ID, 상태, 수금예정일, 미수금, 미수일수
        'date_cols': [1, 12],                  # 기록일, 납품일자 (수금예정일 N은 수식이라 제외)
    },
    '2.매입': {
        'header_row': 1,
        'max_col': 21,
        'formula_cols': [2, 3, 14, 19, 20],   # 매입ID, 상태, 결제예정일, 미지급금, 지연일수
        'date_cols': [1, 12],                  # 기록일, 입고일자
    },
    '3.재고': {
        'header_row': 1,
        'max_col': 12,
        'formula_cols': [8, 11],               # 재고가치, 회전상태
        'date_cols': [9, 10],                  # 마지막입고일, 마지막출고일
    },
    '4.영수증': {
        'header_row': 4,
        'max_col': 8,                          # A~H만 데이터 (J~L은 합계 영역)
        'formula_cols': [2],                   # 영수증ID
        'date_cols': [1],                      # 날짜
    },
    '5.거래처': {
        'header_row': 1,
        'max_col': 12,
        'formula_cols': [10, 11],              # 미수금, 누적매출
        'date_cols': [7],                      # 거래시작일
    },
    '6.통장': {
        'header_row': 1,
        'max_col': 10,                         # A~J만 (L~O는 잔고 요약 영역)
        'formula_cols': [7],                   # 누적잔고
        'date_cols': [1],                      # 일자
    },
    '7.정기업무': {
        'header_row': 4,
        'max_col': 9,
        'formula_cols': [],
        'date_cols': [4, 5],                   # 마지막 실행일, 다음 실행일
    },
    '8.미수관리': {
        'header_row': 4,
        'max_col': 10,
        'formula_cols': [2, 3, 4, 5, 6, 7, 8],  # 모든 집계는 수식
        'date_cols': [9],                       # 마지막 독촉일
    },
    '0.개선아이디어': {
        'header_row': 4,
        'max_col': 6,
        'formula_cols': [],
        'date_cols': [1],
    },
}

# 기본 경로
DEFAULT_XLSX = 'workbook/output/SL철강_5월_운영시트.xlsx'
DATA_DIR = Path('data')


# ============================================================
# 값 변환 헬퍼
# ============================================================
def to_csv_value(value, is_date_col):
    """xlsx 값 → CSV에 쓸 문자열"""
    if value is None:
        return ''
    if isinstance(value, (datetime, date)):
        return value.strftime('%Y-%m-%d')
    return str(value)


def from_csv_value(s, is_date_col):
    """CSV 문자열 → xlsx에 넣을 값"""
    if s == '' or s is None:
        return None
    if is_date_col:
        # 날짜 파싱 시도
        try:
            return datetime.strptime(s, '%Y-%m-%d').date()
        except ValueError:
            pass
        try:
            return datetime.strptime(s, '%Y-%m-%d %H:%M:%S')
        except ValueError:
            pass
        return s  # 파싱 실패 시 문자열 그대로
    # 숫자 시도
    try:
        if '.' in s:
            return float(s)
        return int(s)
    except ValueError:
        pass
    return s


def get_headers(ws, header_row, max_col):
    """헤더 행에서 컬럼명 추출"""
    return [ws.cell(row=header_row, column=c).value for c in range(1, max_col + 1)]


# ============================================================
# extract: xlsx → CSV
# ============================================================
def extract(xlsx_path: str = DEFAULT_XLSX, data_dir: Path = DATA_DIR):
    """현재 xlsx의 입력 데이터를 CSV로 추출 (자동수식 컬럼은 제외)"""
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        print(f'✗ xlsx not found: {xlsx_path}')
        sys.exit(1)

    wb = load_workbook(xlsx_path, data_only=True)
    data_dir.mkdir(parents=True, exist_ok=True)

    total_rows = 0
    for sheet_name, meta in SHEETS.items():
        if sheet_name not in wb.sheetnames:
            print(f'⚠ skip {sheet_name} (시트 없음)')
            continue
        ws = wb[sheet_name]
        header_row = meta['header_row']
        max_col = meta['max_col']
        date_cols = set(meta['date_cols'])

        headers = get_headers(ws, header_row, max_col)
        rows = []
        for r in range(header_row + 1, ws.max_row + 1):
            row_values = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
            # 첫 번째 데이터 컬럼이 비어있으면 빈 행으로 간주
            # 단, 매출/매입은 1열이 기록일이라 그 기준
            # 거래처/재고/통장/영수증/미수관리/개선아이디어/정기업무도 1열 기준 OK
            if row_values[0] is None or row_values[0] == '':
                continue
            rows.append(row_values)

        csv_path = data_dir / f'{sheet_name}.csv'
        with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            for row in rows:
                csv_row = []
                for col_idx, value in enumerate(row, 1):
                    is_date = col_idx in date_cols
                    csv_row.append(to_csv_value(value, is_date))
                writer.writerow(csv_row)

        print(f'✓ {sheet_name:18} → {csv_path.name:25} ({len(rows)}행)')
        total_rows += len(rows)

    print(f'\n총 {total_rows}행 추출 완료. data/ 디렉토리 확인.')


# ============================================================
# rebuild: 빈 양식 + CSV → xlsx
# ============================================================
def rebuild(xlsx_path: str = DEFAULT_XLSX, data_dir: Path = DATA_DIR):
    """빈 양식 새로 생성하고 CSV 데이터를 주입"""
    xlsx_path = Path(xlsx_path)

    # 1. 빈 양식 생성 (with_samples=False)
    print(f'1) 빈 양식 생성: {xlsx_path}')
    build_workbook(str(xlsx_path), with_samples=False)

    # 2. CSV 데이터 주입
    if not data_dir.exists():
        print(f'⚠ {data_dir}/ 없음 — 빈 양식 그대로 저장')
        return

    wb = load_workbook(xlsx_path)
    total_rows = 0

    for sheet_name, meta in SHEETS.items():
        csv_path = data_dir / f'{sheet_name}.csv'
        if not csv_path.exists():
            print(f'⚠ skip {sheet_name} (CSV 없음)')
            continue

        ws = wb[sheet_name]
        header_row = meta['header_row']
        max_col = meta['max_col']
        formula_cols = set(meta['formula_cols'])
        date_cols = set(meta['date_cols'])

        with open(csv_path, encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            csv_headers = next(reader)  # 헤더 스킵
            row_count = 0
            for i, row in enumerate(reader):
                target_row = header_row + 1 + i
                for col_idx, value_str in enumerate(row, 1):
                    if col_idx > max_col:
                        break
                    if col_idx in formula_cols:
                        continue  # 수식 컬럼은 건드리지 않음
                    is_date = col_idx in date_cols
                    parsed = from_csv_value(value_str, is_date)
                    if parsed is not None:
                        ws.cell(row=target_row, column=col_idx, value=parsed)
                row_count += 1

        print(f'✓ {sheet_name:18} ← {csv_path.name:25} ({row_count}행)')
        total_rows += row_count

    wb.save(xlsx_path)
    print(f'\n총 {total_rows}행 주입 완료. {xlsx_path} 저장됨.')


# ============================================================
# sync: extract + rebuild (양식 변경 시 한 사이클)
# ============================================================
def sync(xlsx_path: str = DEFAULT_XLSX, data_dir: Path = DATA_DIR):
    """현재 xlsx에서 데이터 추출 → 새 양식으로 재빌드"""
    print('=' * 60)
    print('SYNC: 데이터 추출 → 빈 양식 재생성 → 데이터 주입')
    print('=' * 60)
    print('\n[1/2] EXTRACT')
    print('-' * 60)
    extract(xlsx_path, data_dir)
    print('\n[2/2] REBUILD')
    print('-' * 60)
    rebuild(xlsx_path, data_dir)
    print('\n✓ SYNC 완료')


# ============================================================
# 진입점
# ============================================================
def parse_args():
    args = sys.argv[1:]
    cmd = 'sync'
    xlsx_path = DEFAULT_XLSX
    data_dir = DATA_DIR
    i = 0
    while i < len(args):
        a = args[i]
        if a in ('extract', 'rebuild', 'sync'):
            cmd = a
        elif a == '--xlsx':
            xlsx_path = args[i + 1]
            i += 1
        elif a == '--data':
            data_dir = Path(args[i + 1])
            i += 1
        elif a in ('-h', '--help'):
            print(__doc__)
            sys.exit(0)
        i += 1
    return cmd, xlsx_path, data_dir


if __name__ == '__main__':
    cmd, xlsx_path, data_dir = parse_args()
    if cmd == 'extract':
        extract(xlsx_path, data_dir)
    elif cmd == 'rebuild':
        rebuild(xlsx_path, data_dir)
    elif cmd == 'sync':
        sync(xlsx_path, data_dir)

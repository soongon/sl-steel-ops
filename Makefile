# SH철강 운영 시스템 Makefile
#
# 일반적인 운영 명령:
#   make build      - 첫 워크북 생성 (샘플 포함)
#   make extract    - xlsx → CSV
#   make rebuild    - CSV → xlsx (빈 양식 + 데이터 주입)
#   make sync       - extract + rebuild (양식 변경 사이클)
#   make backup     - data/를 backup/YYYY-MM-DD/로 복사
#   make clean      - 생성물 삭제 (CSV는 보존)

.PHONY: install build extract rebuild sync backup clean help

XLSX = workbook/output/SL철강_5월_운영시트.xlsx
TODAY := $(shell date +%Y-%m-%d)

help:
	@echo "SH철강 운영 시스템 명령어:"
	@echo ""
	@echo "  make install   - 의존성 설치 (openpyxl)"
	@echo "  make build     - 첫 워크북 생성 (샘플 포함, 5월 1일 첫 시작용 - 이후엔 사용 X)"
	@echo "  make extract   - xlsx → CSV (운영 데이터 추출)"
	@echo "  make rebuild   - CSV → xlsx (빈 양식 + 데이터 주입, 양식 변경 후 권장)"
	@echo "  make sync      - extract + rebuild (매주 일요일 또는 양식 변경 시)"
	@echo "  make backup    - data/를 backup/$(TODAY)/로 복사"
	@echo "  make clean     - 생성된 xlsx, __pycache__ 삭제 (CSV는 보존)"
	@echo ""
	@echo "⚠ 주의: make build는 데이터 있는 상태에서 호출하면 안전장치가 거부합니다."
	@echo "  (실데이터 손실 사고 방지). 양식만 다시 만들고 싶으면 'make rebuild'."

install:
	pip install openpyxl

# build: ⚠ 첫 시작 (5/1) 1회만. 이후엔 데이터가 샘플로 덮어씌워질 수 있어 안전장치가 거부.
#        양식·데이터 모두 다시 만들고 싶으면 'make rebuild' (CSV 데이터 보존).
build:
	python3 workbook/create_spreadsheets.py
	@echo ""
	@echo "✓ 워크북 생성됨: $(XLSX)"
	@echo "→ iCloud Drive 또는 Google Drive에 업로드해서 친구·직원과 공유"

extract:
	python3 migrate.py extract

rebuild:
	python3 migrate.py rebuild

sync:
	python3 migrate.py sync
	@echo ""
	@echo "다음 단계: git add data/ && git commit -m 'data: $(TODAY) 동기화'"

backup:
	@mkdir -p backup/$(TODAY)
	@cp -r data/* backup/$(TODAY)/ 2>/dev/null || true
	@cp $(XLSX) backup/$(TODAY)/ 2>/dev/null || true
	@echo "✓ backup/$(TODAY)/ 에 백업 완료"
	@ls -la backup/$(TODAY)/

clean:
	rm -rf workbook/__pycache__
	rm -f workbook/output/*.xlsx
	@echo "✓ 생성물 삭제 완료 (data/ CSV는 보존됨)"

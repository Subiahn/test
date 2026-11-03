# netscout-auto v0.89 설계

## 목표

이전 릴리스에서 실제 장비로 검증한 조회·후처리·Hostname·Application 등록 기능을 정제된 릴리스로 제공한다.

## 구조

- `python -m netscout`: Inventory, Top N, Detail, 후보 추출, Hostname 조회
- `tools/registration/make_registration.py`: Hostname 결과를 등록 검토 Excel로 변환
- `tools/registration/netscout_sync.py`: 등록 계획 및 명시적 승인 후 적용
- `docker/hostname-lab`: Hostname 조회용 로컬 HTTPS 테스트 환경
- `tools/nca`: 향후 NCA 연동 모듈 위치

## 등록 안전 원칙

- 기본 실행은 GET 기반 Plan-only다.
- 실제 적용은 `--apply`와 `SYNC` 확인을 모두 요구한다.
- `CONFLICT`가 있으면 전체 적용을 중단한다.
- CREATE와 기존 IP/Port 추가만 지원한다.
- 삭제, 기존 IP/Port 제거, 모호한 이름 변경은 자동화하지 않는다.
- 적용 후 GET으로 재검증하고 동일 입력은 `MATCHED`가 되어야 한다.

## 배포 제외 항목

- `.env.local`
- `config/config.test.yaml`
- `config/hostname_lookup.yaml`
- `output`, `results`, `validation_runs`의 실행 산출물
- Python 및 pytest 캐시

## 검증 기준

- 전체 pytest 통과
- compileall 통과
- 메인 CLI와 등록 CLI 도움말 실행
- Docker Hostname Lab 구성 유효성 확인
- 민감정보 및 구버전 문자열 검사

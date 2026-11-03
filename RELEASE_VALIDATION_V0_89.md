# netscout-auto v0.89 릴리스 검증 결과

검증일: 2026-11-03

## 결과

- pytest: `31 passed, 4 subtests passed`
- compileall: PASS
- `python -m netscout --help`: PASS
- `tools/registration/netscout_sync.py --help`: PASS
- Docker Hostname Lab Compose 구문: PASS
- 실제 환경 파일 제외: PASS
- 실제 테스트 서버/장비/비밀번호 참조 검사: PASS
- 구버전 런타임 문자열 검사: PASS

## 릴리스 포함

- 조회·분석·후처리 CLI
- Hostname 조회와 Docker Lab
- Application 등록 Excel 변환 및 동기화 도구
- 등록 단위·통합·시나리오 테스트
- Windows/Ubuntu/운영/E2E 문서
- 설정 예제와 vendored 의존성

## 릴리스 제외

- `.env.local`
- `config/config.test.yaml`
- `config/hostname_lookup.yaml`
- 실제 장비 대상 YAML
- `output`, `results`, `validation_runs` 실행 결과
- Python/pytest 캐시

## 실제 장비 검증 계보

이 릴리스의 등록 로직은 이전 검증에서 CREATE → VERIFIED → MATCHED와 ADD_IP/ADD_PORT 흐름을 확인했다. v0.89 릴리스 폴더에는 해당 고객/테스트 환경 산출물을 포함하지 않는다.

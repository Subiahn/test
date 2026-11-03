# netscout-auto v0.89 릴리스 노트

## 포함 기능

- Inventory, Top N, Detail 및 Batch 조회
- 등록 Application 비교 및 Detail 대상 제외
- IP/Port 후보 추출과 Hostname 조회
- Registration Excel 생성
- CREATE, MATCHED, ADD_IP, ADD_PORT, 복합 추가 및 CONFLICT 계획
- `--apply` + `SYNC` 승인 후 POST/PUT 및 GET 재검증
- Docker 기반 Hostname Lab

## v0.89 정제 내용

- 런타임 버전 표기를 v0.89로 통일했다.
- 실제 환경 설정과 생성 산출물을 릴리스에서 제거했다.
- 설정 파일 역할과 E2E 절차를 문서화했다.
- PUT 시 전체 desired payload를 유지하도록 보완된 동기화 로직을 포함했다.
- 실제 테스트 장비에서 CREATE → VERIFIED → MATCHED 흐름을 검증했다.

## 제한사항

- Application 삭제와 기존 IP/Port 제거는 지원하지 않는다.
- 등록 Excel은 사람이 검토하고 `Register=Y`를 지정해야 한다.
- 자체 서명 인증서는 테스트에서만 `--insecure`를 사용하고 운영에서는 CA 파일 사용을 권장한다.

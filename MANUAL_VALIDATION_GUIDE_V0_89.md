# netscout-auto v0.89 수동 검증 가이드

이 문서는 운영자가 v0.89의 전체 기능을 직접 실행하고, 각 결과 파일을 눈으로 확인하기 위한 절차다. 조회 단계는 장비 설정을 변경하지 않는다. 실제 변경은 등록 단계에서 `--apply`를 사용하고 `SYNC`를 입력했을 때만 발생한다.

## 1. 검증 범위

```text
환경 및 회귀 검사
  → Inventory BASELINE
  → Top N
  → Detail 대상 생성
  → Pure/Current Inventory 비교 및 화이트리스트 생성
  → 화이트리스트로 Detail 대상 제외
  → Detail Dry-run 및 실제 조회
  → HTTPS/TCP_OTHER 후보 추출
  → Docker Hostname Lab
  → Registration 검토 Excel 생성
  → 사람이 등록값 입력
  → Plan-only 검증
  → 실제 등록 승인
  → Inventory AFTER
  → 동일 파일 재실행(MATCHED)
```

## 2. 안전 원칙

1. 테스트 서버에서만 실행한다.
2. 조회용 비밀번호는 Git에 포함되지 않는 `config\.env.local`에 입력하고, 등록용 비밀번호는 실행 시 프롬프트에 입력한다.
3. `netscout_sync.py`는 먼저 `--apply` 없이 실행한다.
4. 계획에 `CONFLICT`가 있으면 실제 적용하지 않는다.
5. 계획과 Excel이 일치할 때만 `SYNC`를 입력한다.
6. 테스트 IP와 Port는 기존 Application과 겹치지 않는 값으로 승인받는다.
7. 이 도구는 Application 삭제와 기존 IP/Port 제거를 지원하지 않는다.

## 3. 준비

PowerShell을 열고 프로젝트로 이동한다.

```powershell
cd "D:\yury\netscout-api-automation\netscout-auto_v0.89"
```

서버와 API 설정 파일을 준비하고 직접 편집한다.

```powershell
Copy-Item .\config\config.example.yaml .\config\config.test.yaml
notepad .\config\config.test.yaml
```

최소 확인 항목:

```yaml
netscout:
  base_url: "https://the-se.iptime.org:8443"
  auth_type: "basic"
  verify_ssl: false  # 자체 서명 인증서를 쓰는 테스트 환경만
  timeout: 30
```

사내 CA가 있으면 `verify_ssl: true`와 CA 신뢰 구성을 사용한다.

계정값은 환경변수로 설정하지 않고 프로젝트의 로컬 설정 파일에 직접 입력한다.

```powershell
notepad .\config\.env.local
```

```dotenv
NETSCOUT_USERNAME=administrator
NETSCOUT_PASSWORD=실제비밀번호
```

`config\.env.local`은 Git에 커밋하지 않는다. 아래 조회 명령은 항상 `--config`, `--env`, `--output-dir`을 직접 지정한다. 등록 도구는 이 파일을 자동으로 읽지 않으므로 URL과 사용자 이름을 명령에 직접 지정하고, 비밀번호는 실행 시 프롬프트에 입력한다.

문서의 `YYYYMMDD`는 실행일 폴더명으로 직접 바꾼다. 예를 들어 2026년 11월 3일 실행 결과는 `20261103`을 사용한다.

## 4. 단계 0 — 코드와 연결 검사

```powershell
python --version
python -m pytest -q
python -m compileall -q .\V8 .\src .\tools\registration
python -m netscout --help
python .\tools\registration\netscout_sync.py --help
Test-NetConnection the-se.iptime.org -Port 8443
```

통과 기준:

- Python 3.11 이상
- pytest 실패 0건
- compileall 오류 없음
- CLI 도움말에 `v0.89` 표시
- `TcpTestSucceeded=True`

## 5. 단계 1 — Inventory BASELINE

```powershell
python -m netscout inventory `
  --config .\config\config.test.yaml `
  --env .\config\.env.local `
  --output-dir .\output `
  --output-file "Inventory_BASELINE_YYYYMMDD.xlsx"
```

결과 위치:

```text
output/<날짜>/00_inventory/Inventory_BASELINE_<날짜>.xlsx
```

Excel 확인:

- `Application_List`, `Device_List`, `Interface_List`, `Run_Summary` 시트 존재
- Application/Device/Interface 행 수가 로그와 일치
- 사용할 Device Name과 Interface Name/Number 기록
- 테스트에 사용할 Application 이름이 이미 존재하는지 확인

기록:

```text
Application 수:
Device:
Interface:
테스트 앱 기존 존재 여부:
```

## 6. 단계 2 — Top N

Inventory에서 확인한 실제 Device와 Interface를 사용한다.

```powershell
python -m netscout top `
  --config .\config\config.test.yaml `
  --env .\config\.env.local `
  --device "nG1vstream635" `
  --interface "if3" `
  --duration LAST_7_DAYS `
  --top-n 50 `
  --output-dir .\output
```

통과 기준:

- `01_top`에 Excel 생성
- Rank가 1부터 연속
- Bit Rate가 내림차순
- 행 수가 50 이하
- Device/Interface가 입력값과 일치

## 7. 단계 3 — Detail 대상 생성

`<TopN파일>`을 실제 생성 파일명으로 바꾼다.

```powershell
python -m netscout detail-targets-yaml `
  --top-xlsx ".\output\YYYYMMDD\01_top\<TopN파일>.xlsx" `
  --detail-top-n 50 `
  --limit 1000
```

결과:

```text
output/<날짜>/02_targets/detail_targets.yaml
```

YAML 확인:

- 대상 수 50 이하
- `rank`, `application`, `device`, `interface`, `duration` 존재
- 조회하지 않을 Application은 `enabled: false`로 변경 가능

### 단계 3-1 — 수동 등록 Application 화이트리스트 비교 및 제외

앱 등록 전 보관한 순정 Inventory와 현재 BASELINE Inventory를 비교한다. 순정 파일은 `pure_xlsx\pure_inventory.xlsx`에 보관해야 한다.

```powershell
python -m netscout compare-inventory-apps `
  --pure-inventory ".\pure_xlsx\pure_inventory.xlsx" `
  --current-inventory ".\output\YYYYMMDD\00_inventory\Inventory_BASELINE_YYYYMMDD.xlsx" `
  --output ".\output\YYYYMMDD\02_targets\manual_registered_apps.csv" `
  --yaml-output ".\output\YYYYMMDD\02_targets\manual_registered_apps.yaml" `
  --sheet Application_List `
  --key-column ApplicationTypeCode
```

`PURE`에는 없고 `CURRENT`에만 있는 `ApplicationTypeCode`를 수동 등록 Application으로 판정한다. 생성된 CSV와 YAML을 검토한 뒤 Detail 대상에 반영한다.

```powershell
python -m netscout filter-detail-targets `
  --top-yaml ".\output\YYYYMMDD\02_targets\detail_targets.yaml" `
  --manual-apps-yaml ".\output\YYYYMMDD\02_targets\manual_registered_apps.yaml" `
  --output ".\output\YYYYMMDD\02_targets\detail_targets_filtered.yaml" `
  --dry-run

python -m netscout filter-detail-targets `
  --top-yaml ".\output\YYYYMMDD\02_targets\detail_targets.yaml" `
  --manual-apps-yaml ".\output\YYYYMMDD\02_targets\manual_registered_apps.yaml" `
  --output ".\output\YYYYMMDD\02_targets\detail_targets_filtered.yaml"
```

통과 기준:

- 화이트리스트와 일치하는 대상만 `enabled: false`
- 나머지 대상의 순서와 설정은 변경되지 않음
- 이후 Detail 단계는 반드시 `detail_targets_filtered.yaml` 사용

## 8. 단계 4 — Detail Dry-run

```powershell
python -m netscout run-detail-targets `
  --config .\config\config.test.yaml `
  --env .\config\.env.local `
  --targets ".\output\YYYYMMDD\02_targets\detail_targets_filtered.yaml" `
  --output-dir .\output `
  --default-duration LAST_7_DAYS `
  --dry-run
```

통과 기준:

- 실행 대상이 모두 `planned`
- Device/Interface/Application이 예상값과 일치
- API 데이터 조회와 Excel 생성은 아직 수행되지 않음

## 9. 단계 5 — Detail 실제 조회

```powershell
python -m netscout run-detail-targets `
  --config .\config\config.test.yaml `
  --env .\config\.env.local `
  --targets ".\output\YYYYMMDD\02_targets\detail_targets_filtered.yaml" `
  --output-dir .\output `
  --default-duration LAST_7_DAYS
```

결과:

```text
output/<날짜>/03_detail/<Application>/*.xlsx
```

통과 기준:

- 각 활성 Application 디렉터리에 Excel 생성
- 최종 로그의 `failed=0`
- Excel의 Application, Device, Interface, 기간이 요청값과 일치

실행 제한으로 중단되면 완료된 Application을 `enabled: false`로 설정한 별도 재개 YAML을 만들고 남은 대상만 실행한다.

## 10. 단계 6 — IP/Port 후보 추출

HTTPS:

```powershell
python -m netscout extract-candidates `
  --detail-root ".\output\YYYYMMDD\03_detail\HTTPS" `
  --targets-yaml ".\output\YYYYMMDD\02_targets\detail_targets_filtered.yaml" `
  --output ".\output\YYYYMMDD\04_candidates\HTTPS\ip_port_candidates.xlsx" `
  --yaml-output ".\output\YYYYMMDD\04_candidates\HTTPS\ip_port_candidates.yaml" `
  --applications HTTPS
```

TCP_OTHER:

```powershell
python -m netscout extract-candidates `
  --detail-root ".\output\YYYYMMDD\03_detail\TCP_OTHER" `
  --targets-yaml ".\output\YYYYMMDD\02_targets\detail_targets_filtered.yaml" `
  --output ".\output\YYYYMMDD\04_candidates\TCP_OTHER\ip_port_candidates.xlsx" `
  --yaml-output ".\output\YYYYMMDD\04_candidates\TCP_OTHER\ip_port_candidates.yaml" `
  --applications TCP_OTHER
```

통과 기준:

- Excel과 YAML 모두 생성
- IP/Port/Application이 원본 Detail과 일치
- 중복 제거 전후 수량이 로그에 표시
- `unknown_port_rows`를 확인하고 필요하면 수동 검토

## 11. 단계 7 — Docker Hostname Lab

```powershell
Copy-Item .\config\hostname_lookup.lab.yaml .\config\hostname_lookup.yaml
docker compose -f .\docker\hostname-lab\docker-compose.yml up -d --build
docker ps --filter "name=netscout-hostname-lab"
```

샘플 검증:

```powershell
python -m netscout hostname-lookup `
  --input .\docker\hostname-lab\sample_ip_port_candidates.yaml `
  --config .\config\hostname_lookup.yaml
```

실제 HTTPS 후보 검증:

```powershell
python -m netscout hostname-lookup `
  --input ".\output\YYYYMMDD\04_candidates\HTTPS\ip_port_candidates.yaml" `
  --config .\config\hostname_lookup.yaml
```

결과:

```text
output/<날짜>/05_hostname/HTTPS/hostname_lookup.xlsx
```

통과 기준:

- Docker 포트 5000/5002 실행
- 전체 후보 IP가 처리됨
- Lab 데이터에 있는 IP는 Hostname이 표시됨
- Lab 데이터에 없는 IP는 빈 Hostname 또는 미매칭 상태로 표시됨

## 12. 단계 8 — Registration 검토 Excel

```powershell
python .\tools\registration\make_registration.py `
  --input ".\output\YYYYMMDD\05_hostname\HTTPS\hostname_lookup.xlsx" `
  --output ".\output\YYYYMMDD\05_hostname\HTTPS\registration_review.xlsx"
```

원본을 보존하고 승인본을 따로 만든다.

```powershell
Copy-Item `
  ".\output\YYYYMMDD\05_hostname\HTTPS\registration_review.xlsx" `
  ".\validation_runs\registration_approved_YYYYMMDD.xlsx"
```

## 13. Registration Excel 수동 입력

등록할 행에 다음 값을 직접 입력한다.

| 열 | 예시 | 확인 사항 |
|---|---|---|
| Register | `Y` | 등록할 행만 Y |
| ApplicationName | `AUTO_V089_MANUAL_01` | 최대 32자, 기존 이름과 중복 금지 |
| RegistrationType | `SBA` | IP/Port 기반 서버 앱 |
| ParentApplication | `TCP` | TCP/UDP/SCTP 또는 기존 앱 |
| Protocol | `TCP` | 부모 프로토콜과 일치 |
| IP | 승인된 테스트 IP | 기존 앱 주소와 중복 금지 |
| Port | 승인된 테스트 Port | 1~65535 |
| Description | 수동 E2E 테스트 | 선택값 |

일반적인 TCP 서버 앱 예시:

```text
Register=Y
RegistrationType=SBA
ParentApplication=TCP
Protocol=TCP
```

주의:

- 같은 ApplicationName 행은 IP/Port가 하나의 Application으로 병합될 수 있다.
- 다른 Application 이름인데 IP 범위가 겹치면 `CONFLICT`가 발생해야 정상이다.
- `Register=N` 행은 등록 대상이 아니다.

## 14. 단계 9 — 등록 계획만 확인

먼저 `--apply` 없이 실행한다.

```powershell
python .\tools\registration\netscout_sync.py `
  --input ".\validation_runs\registration_approved_YYYYMMDD.xlsx" `
  --base-url "https://the-se.iptime.org:8443" `
  --username "administrator" `
  --insecure
```

명령 실행 후 Password 프롬프트가 나타나면 실제 비밀번호를 직접 입력한다. 입력값은 화면에 표시되지 않는다.

예상 Action:

| Action | 의미 | 적용 여부 |
|---|---|---|
| CREATE | 이름이 없는 신규 앱 | 승인 후 적용 가능 |
| MATCHED | 현재 설정과 동일 | 변경 없음 |
| ADD_IP | 기존 앱에 IP 추가 | 승인 후 적용 가능 |
| ADD_PORT | 기존 앱에 Port 추가 | 승인 후 적용 가능 |
| ADD_IP+ADD_PORT | IP와 Port 추가 | 승인 후 적용 가능 |
| CONFLICT | 이름·주소·유형 충돌 | 적용 금지 |

통과 기준:

- Excel과 계획의 이름/IP/Port가 일치
- 의도하지 않은 Application이 없음
- `CONFLICT`가 없음
- `--apply`가 없으므로 POST/PUT이 실행되지 않음
- `results/sync_<시각>/plan.json` 생성

## 15. 단계 10 — 실제 등록

계획이 정확하고 테스트 장비 변경 승인을 받은 경우에만 실행한다.

```powershell
python .\tools\registration\netscout_sync.py `
  --input ".\validation_runs\registration_approved_YYYYMMDD.xlsx" `
  --base-url "https://the-se.iptime.org:8443" `
  --username "administrator" `
  --insecure `
  --apply
```

화면에 계획이 다시 표시되면 확인 후 다음을 정확히 입력한다.

```text
SYNC
```

통과 기준:

- `Applied and verified N change(s)` 표시
- `results.json`의 상태가 `VERIFIED`
- CREATE/PUT XML에 승인한 이름, IP, Port가 기록됨
- 예상과 다르면 `SYNC`를 입력하지 않고 취소

## 16. 단계 11 — Inventory AFTER

```powershell
python -m netscout inventory `
  --config .\config\config.test.yaml `
  --env .\config\.env.local `
  --output-dir .\output `
  --output-file "Inventory_AFTER_YYYYMMDD.xlsx"
```

Excel에서 다음을 확인한다.

- `ApplicationShortName`이 승인한 ApplicationName과 일치
- ApplicationType이 `Server-based Apps`
- ParentApplication이 예상값과 일치
- IsActive가 `true`
- Application 수가 CREATE 개수만큼 증가

서버 IP와 Port의 최종 근거는 등록 결과 XML과 적용 후 GET 검증 결과를 함께 확인한다. Inventory의 `ApplicationPort`는 서버 서비스 Port가 아니라 장비 내부 Application 식별값일 수 있다.

## 17. 단계 12 — 멱등성 검증

적용에 사용한 동일 Excel을 `--apply` 없이 다시 실행한다.

```powershell
python .\tools\registration\netscout_sync.py `
  --input ".\validation_runs\registration_approved_YYYYMMDD.xlsx" `
  --base-url "https://the-se.iptime.org:8443" `
  --username "administrator" `
  --insecure
```

통과 기준:

```text
MATCHED
Plan only; 0 applicable change(s)
```

## 18. 오류 시 확인 순서

### Missing username/password

```powershell
Get-Content .\config\.env.local | Select-String '^NETSCOUT_USERNAME='
notepad .\config\.env.local
```

- 조회 명령이면 `config\.env.local`의 사용자명과 비밀번호를 확인한다.
- 등록 명령이면 `--username "administrator"`가 포함됐는지 확인하고 Password 프롬프트에 비밀번호를 직접 입력한다.
- 비밀번호 확인을 위해 `Get-Content`로 파일 전체를 콘솔에 출력하지 않는다.

### Config API connection failed

- `config.test.yaml`의 `base_url` 확인
- 포트 포함 여부 확인
- `Test-NetConnection` 실행
- 테스트 환경에서 인증서 설정 확인

### CONFLICT

- 실제 적용 금지
- `plan.json`의 충돌 대상 확인
- ApplicationName, IP/CIDR, Port를 기존 Inventory와 비교

### Detail 실행 제한시간 초과

- 이미 생성된 Application Excel을 확인
- 완료된 대상을 비활성화한 재개 YAML 사용
- 기존 결과를 삭제하지 않고 남은 대상만 실행

## 19. 최종 결과 기록표

| 항목 | 결과 | 증적 위치 |
|---|---|---|
| pytest | PASS/FAIL | 콘솔 로그 |
| Inventory BASELINE | PASS/FAIL | `00_inventory` |
| Top N | PASS/FAIL | `01_top` |
| Detail 대상 | PASS/FAIL | `02_targets` |
| Detail Batch | PASS/FAIL | `03_detail` |
| 후보 추출 | PASS/FAIL | `04_candidates` |
| Hostname | PASS/FAIL | `05_hostname` |
| Registration Excel | PASS/FAIL | 승인 Excel |
| Plan-only | PASS/FAIL | `results/.../plan.json` |
| 실제 적용 | PASS/FAIL/SKIP | `results/.../results.json` |
| Inventory AFTER | PASS/FAIL/SKIP | `00_inventory` |
| MATCHED 재실행 | PASS/FAIL/SKIP | `results/.../plan.json` |

## 20. 최종 PASS 기준

다음을 모두 만족하면 전체 수동 검증 PASS다.

1. 자동 테스트와 컴파일이 통과한다.
2. Inventory, Top N, Detail 결과가 요청값과 일치한다.
3. 후보 Excel/YAML이 원본 Detail과 추적 가능하다.
4. Hostname 조회 성공과 미매칭을 구분할 수 있다.
5. Registration Excel에서 사람이 승인한 행만 계획에 포함된다.
6. `CONFLICT`가 있으면 쓰기가 차단된다.
7. `--apply` 전에는 장비 쓰기가 없다.
8. 실제 적용 결과가 `VERIFIED`다.
9. AFTER Inventory에서 대상 Application이 확인된다.
10. 동일 입력 재실행 결과가 `MATCHED`이고 변경 0건이다.

## 21. 테스트 종료

Docker Lab만 종료하려면 다음을 실행한다.

```powershell
docker compose -f .\docker\hostname-lab\docker-compose.yml down
```

실제 장비에 생성한 테스트 Application은 이 도구로 자동 삭제하지 않는다. 삭제가 필요하면 NETSCOUT UI 또는 승인된 별도 운영 절차로 대상 이름과 Type Code를 재확인한 후 처리한다.

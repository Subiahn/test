# NETSCOUT Automation v0.89 전체 E2E 검증 시나리오

## 1. 검증 목적

조회·분석·후처리 흐름과 독립 Application 등록 도구를 v0.89 릴리스 기준으로 처음부터 끝까지 연결해 검증한다.

```text
환경/코드 검사
  → 등록 전 Inventory(PURE)
  → Top N
  → Detail 대상 YAML
  → 수동 등록 Application 비교/제외
  → Detail Batch
  → IP/Port 후보 Excel·YAML
  → Hostname 조회 Excel
  → Registration Excel 변환·사람 검토
  → 등록 전 Inventory(BEFORE)
  → 등록 계획(CREATE/MATCHED/ADD/CONFLICT)
  → 승인된 테스트 Application 적용
  → 등록 후 Inventory(AFTER)
  → Excel 전후 비교
  → 동일 파일 재실행(MATCHED)
```

## 2. 검증 범위

| 영역 | 검증 내용 |
|---|---|
| 기존 CLI | Inventory, Top N, Detail 대상, Detail Batch, 후보 추출, Hostname 조회 |
| 후처리 | Pure/Current Inventory 비교, 수동 등록 앱 제외, IP/Port 중복 제거 |
| 등록 Excel | Hostname 결과 변환, 수동 검토 필드, 동일 이름 행 병합 |
| 등록 계획 | CREATE, MATCHED, ADD_IP, ADD_PORT, ADD_IP+ADD_PORT, CONFLICT |
| 등록 적용 | 명시적 `--apply`와 `SYNC`, POST/PUT, 적용 후 GET 검증 |
| 안전성 | 기본 읽기 전용, 충돌 시 전체 중단, 삭제 및 기존 값 제거 금지 |
| 회귀 | 기존 조회 명령 유지, 등록 도구가 별도 `tools/registration`에 존재 |

운영 서버는 범위 밖이다. 실제 쓰기는 승인된 테스트 서버와 테스트 전용 Application/IP/Port에만 수행한다.

## 3. 환경과 변수

프로젝트 루트:

```powershell
cd "D:\yury\netscout-api-automation\netscout-auto_v0.89"
```

공통 변수 예시:

```powershell
$BASE_URL = "https://test-server.example.com:8443"
$CONFIG = ".\config\config.test.yaml"
$OUTPUT = ".\output"
$RUN_DATE = Get-Date -Format "yyyyMMdd"
$env:NETSCOUT_USERNAME = "dummy-user"
$env:NETSCOUT_PASSWORD = Read-Host "NETSCOUT password"
```

자체 서명 인증서를 쓰는 현재 테스트 환경은 설정의 `verify_ssl: false` 또는 등록 도구의 `--insecure`를 사용한다. 사내 CA를 받으면 `--ca-file`로 교체한다.

## 4. 테스트 데이터 준비

Application 등록 쓰기 테스트용 값은 담당자가 승인한 값으로 채운다.

| 항목 | 테스트 값 | 승인/확인 |
|---|---|---|
| 신규 Application 이름 | `AUTO_V088_<날짜>_01` | |
| RegistrationType | `SBA` | |
| ParentApplication | `TCP` | |
| Protocol | `TCP` | |
| 최초 IP | 승인된 미사용 IP | |
| 추가 IP | 승인된 미사용 IP | |
| 최초 Port | 승인된 미사용 Port | |
| 추가 Port | 승인된 미사용 Port | |
| 삭제 계획 | 없음 | 고정 |

IP는 다른 Application의 주소/CIDR과 겹치지 않아야 한다. 적용 전 반드시 현재 Inventory에서 이름과 주소 중복을 확인한다.

## 5. 단계 0 — 코드와 환경 회귀 검사

```powershell
python --version
python -m pytest -q
python -m compileall -q .\V8 .\src .\tools\registration
python -m netscout --help
python .\tools\registration\netscout_sync.py --help
Test-NetConnection test-server.example.com -Port 8443
```

통과 기준:

- pytest 실패 0건
- compileall 오류 없음
- 기존 CLI 도움말은 v0.89 명령을 유지
- 등록 명령은 기존 CLI가 아니라 `tools/registration`에서만 제공
- `TcpTestSucceeded=True`

## 6. 단계 1 — 등록 전 기준 Inventory(PURE) 생성

아직 테스트 Application을 등록하지 않은 상태에서 실행한다.

```powershell
python -m netscout inventory `
  --config $CONFIG `
  --output-dir $OUTPUT `
  --output-file "Inventory_PURE_$RUN_DATE.xlsx"
```

검증 Excel: `Application_List`, `Device_List`, `Interface_List`, `Run_Summary`

| 확인 항목 | 통과 기준 |
|---|---|
| Application 행 수 | API 로그 개수와 `Application_List` 데이터 행 수 일치 |
| Device 행 수 | API 로그 개수와 `Device_List` 데이터 행 수 일치 |
| Interface 행 수 | API 로그 개수와 `Interface_List` 데이터 행 수 일치 |
| Application 키 | `ApplicationTypeCode` 또는 `ApplicationShortName` 중복 여부 검토 |
| 테스트 앱 | 신규 테스트 이름이 존재하지 않음 |

이 파일은 이후 비교의 기준이므로 수정하지 않고 보관한다.

## 7. 단계 2 — Top N 조회

단계 1의 `Device_List`와 `Interface_List`에서 실제 Device Name과 Interface Name/Number를 선택한다.

```powershell
python -m netscout top `
  --config $CONFIG `
  --device "nG1vstream635" `
  --interface "if3" `
  --duration LAST_7_DAYS `
  --top-n 50 `
  --output-dir $OUTPUT
```

통과 기준:

- Top N Excel 생성
- Rank가 1부터 순서대로 존재
- Bit Rate가 내림차순
- 선택한 Device/Interface/기간이 Run Summary와 일치
- 행 수가 `top-n` 이하

## 8. 단계 3 — Detail 대상 YAML 생성 및 검토

```powershell
python -m netscout detail-targets-yaml `
  --top-xlsx ".\output\$RUN_DATE\01_top\<TopN파일>.xlsx" `
  --output ".\output\$RUN_DATE\02_targets\detail_targets.yaml" `
  --detail-top-n 50 `
  --limit 1000 `
  --dedup-ip
```

Top N Excel과 YAML 비교:

| 비교 키 | Top N Excel | detail_targets.yaml | 통과 기준 |
|---|---|---|---|
| Application | Application 이름 열 | target application | 동일 |
| 순위 | Rank | 대상 순서/선택 범위 | 1~50 범위 |
| Device | 조회 Device | device | 동일 |
| Interface | 조회 Interface | interface | 동일 |
| 활성화 | 해당 없음 | enabled | 기본값 검토 |

사람이 YAML의 `enabled`, duration, limit를 검토한 후 다음 단계로 진행한다.

## 9. 단계 4 — 수동 등록 Application 비교 및 Detail 대상 제외

### 9.1 Pure/Current Inventory 비교

기준 Inventory와 현재 Inventory를 비교해 그 사이 수동 등록된 Application을 찾는다.

```powershell
python -m netscout compare-inventory-apps `
  --pure-inventory ".\output\$RUN_DATE\00_inventory\Inventory_PURE_$RUN_DATE.xlsx" `
  --current-inventory ".\output\$RUN_DATE\00_inventory\Inventory_CURRENT_$RUN_DATE.xlsx" `
  --output ".\output\$RUN_DATE\02_targets\manual_registered_apps.csv" `
  --yaml-output ".\output\$RUN_DATE\02_targets\manual_registered_apps.yaml" `
  --sheet Application_List `
  --key-column ApplicationTypeCode
```

Excel 비교 규칙:

- `PURE.Application_List[ApplicationTypeCode]`에는 없고 `CURRENT`에만 있는 키를 신규 수동 등록 앱으로 판정
- 결과 CSV의 각 행은 CURRENT Excel의 같은 키 행과 필드가 일치해야 함
- 키가 비어 있거나 중복되면 이름과 유형을 추가로 사람이 검토

### 9.2 Detail 대상에서 수동 등록 앱 제외

```powershell
python -m netscout filter-detail-targets `
  --top-yaml ".\output\$RUN_DATE\02_targets\detail_targets.yaml" `
  --manual-apps-yaml ".\output\$RUN_DATE\02_targets\manual_registered_apps.yaml" `
  --output ".\output\$RUN_DATE\02_targets\detail_targets_filtered.yaml" `
  --dry-run

python -m netscout filter-detail-targets `
  --top-yaml ".\output\$RUN_DATE\02_targets\detail_targets.yaml" `
  --manual-apps-yaml ".\output\$RUN_DATE\02_targets\manual_registered_apps.yaml" `
  --output ".\output\$RUN_DATE\02_targets\detail_targets_filtered.yaml"
```

통과 기준: 수동 등록 목록과 일치하는 대상만 `enabled: false`가 되고, 나머지 대상 설정은 변경되지 않는다.

## 10. 단계 5 — Detail Batch

먼저 dry-run으로 호출 대상만 검토한다.

```powershell
python -m netscout run-detail-targets `
  --config $CONFIG `
  --targets ".\output\$RUN_DATE\02_targets\detail_targets_filtered.yaml" `
  --output-dir $OUTPUT `
  --default-duration LAST_7_DAYS `
  --dry-run
```

확인 후 실제 조회:

```powershell
python -m netscout run-detail-targets `
  --config $CONFIG `
  --targets ".\output\$RUN_DATE\02_targets\detail_targets_filtered.yaml" `
  --output-dir $OUTPUT `
  --default-duration LAST_7_DAYS
```

통과 기준:

- `enabled: true` 대상만 API 호출
- Application별 Detail Excel 생성
- Source/Destination IP와 Port가 존재
- `--dedup-ip` 적용 시 중복 조합 제거
- 실패 대상은 로그와 Run Summary에서 식별 가능

## 11. 단계 6 — IP/Port 후보 추출

```powershell
python -m netscout extract-candidates `
  --detail-root ".\output\$RUN_DATE\03_detail" `
  --targets-yaml ".\output\$RUN_DATE\02_targets\detail_targets_filtered.yaml" `
  --output ".\output\$RUN_DATE\04_candidates\ip_port_candidates.xlsx" `
  --yaml-output ".\output\$RUN_DATE\04_candidates\ip_port_candidates.yaml"
```

Detail Excel과 후보 Excel 비교:

| 비교 내용 | 통과 기준 |
|---|---|
| Application | 원본 Detail의 Application과 후보 행 연결 가능 |
| IP/Port | 후보의 IP/Port가 원본 Detail에 존재 |
| 중복 제거 | 동일 Application/IP/Port가 중복되지 않음 |
| 필터 | disabled Application이 기본적으로 제외됨 |
| 추적성 | 후보에서 원본 Application을 식별 가능 |

대표 Application 3개 이상을 골라 원본 Detail과 후보 Excel을 수동 대조한다.

## 12. 단계 7 — Hostname 조회

```powershell
python -m netscout hostname-lookup `
  --input ".\output\$RUN_DATE\04_candidates\ip_port_candidates.yaml" `
  --config ".\config\hostname_lookup.yaml" `
  --output ".\output\$RUN_DATE\05_hostname\hostname_lookup.xlsx"
```

후보 Excel과 Hostname Excel 비교:

- 후보의 모든 조회 대상 IP가 Hostname 결과에 존재하거나 실패 사유가 기록됨
- IP/Port/Application 값이 변형되지 않음
- 성공 행은 Hostname과 Found 상태가 일치
- 실패 행은 Error 또는 미조회 상태를 식별 가능
- 같은 IP가 여러 Application/Port에 속하면 원본 관계가 유지됨

## 13. 단계 8 — Registration Excel 생성 및 사람 검토

```powershell
python .\tools\registration\make_registration.py `
  --input ".\output\$RUN_DATE\05_hostname\hostname_lookup.xlsx" `
  --output ".\output\$RUN_DATE\05_hostname\registration_review.xlsx"
```

Hostname Excel과 Registration Excel 비교:

| Hostname 결과 | Registration 열 | 통과 기준 |
|---|---|---|
| IP | IP | 값 동일 |
| Port | Port | 값 동일 |
| Hostname | Hostname | 값 동일 |
| Application | Application | 원본 값 보존 |
| Found/Error | Found/Error/Review | 실패·주의 상태 보존 |
| Source row | Source Sheet/Source Row | 원본 추적 가능 |

자동 변환 후 사람이 반드시 채우는 열:

- `Register`: 등록할 행만 `Y`
- `ApplicationName`: 최대 32자, 금지문자 제외
- `RegistrationType`: `SBA`, `WKA`, `CHILD`, `URL`
- `ParentApplication`: TCP/UDP/SCTP 또는 기존 Application
- `Protocol`: 유형과 부모 규칙에 맞게 입력
- 필요 시 Description, Group, Tag 및 유형별 옵션

검토 완료본은 원본과 다른 이름으로 저장한다.

## 14. 단계 9 — 유형별 등록 파일 검증

| 유형 | 필수 핵심값 | 검증 기준 |
|---|---|---|
| SBA | 이름, TCP/UDP/SCTP 부모, 같은 Protocol, IP, Port | 서버 주소/포트 기반 앱 |
| WKA | 이름, TCP/UDP/SCTP 부모, 같은 Protocol, 단일 Port | IP 사용 금지 |
| CHILD | 이름, 기존 Application 부모, IP | TCP/UDP/SCTP/IP 직접 부모 금지 |
| URL | 이름, HTTP/HTTPS 부모, URLAddress | URL scheme과 부모 일치, Port/Description 사용 금지 |

동일 Application 이름을 여러 행에 쓰는 경우 주소 외 설정과 포트가 모두 같아야 한다. 같은 설정이면 IP만 병합되고, 다른 포트/설정이면 오류로 중단되어야 한다.

## 15. 단계 10 — 등록 직전 Inventory(BEFORE)와 계획

등록 직전에 최신 Inventory를 다시 저장한다.

```powershell
python -m netscout inventory `
  --config $CONFIG `
  --output-dir $OUTPUT `
  --output-file "Inventory_BEFORE_REGISTER_$RUN_DATE.xlsx"

python .\tools\registration\netscout_sync.py `
  --input ".\output\$RUN_DATE\05_hostname\registration_approved.xlsx" `
  --base-url $BASE_URL `
  --username $env:NETSCOUT_USERNAME `
  --insecure
```

계획 결과 판정:

| Action | 의미 | 다음 조치 |
|---|---|---|
| CREATE | 이름이 없어 신규 생성 가능 | 주소 중복 최종 검토 후 승인 |
| MATCHED | 이미 같은 상태 | 쓰기 불필요 |
| ADD_IP | 기존 앱에 신규 IP 합집합 추가 | 기존 IP 보존 확인 |
| ADD_PORT | 기존 앱에 신규 Port 합집합 추가 | 기존 Port 보존 확인 |
| ADD_IP+ADD_PORT | IP와 Port 모두 추가 | 교차 영향 검토 |
| CONFLICT | 주소, 부모, 유형 등 충돌 | 적용 금지, 입력/설계 재검토 |

`results/sync_<timestamp>/plan.json`의 action/name/message를 승인표에 기록한다.

## 16. 단계 11 — 실제 등록 적용(승인 Gate)

다음 조건을 모두 충족해야 한다.

- 테스트 서버임을 확인
- 테스트 Application 이름/IP/Port 승인 완료
- 최신 BEFORE Inventory 보관
- 계획에 `CONFLICT` 없음
- CREATE/ADD 목록을 사람이 검토
- 삭제 또는 기존 값 제거가 없음을 확인

```powershell
python .\tools\registration\netscout_sync.py `
  --input ".\output\$RUN_DATE\05_hostname\registration_approved.xlsx" `
  --base-url $BASE_URL `
  --username $env:NETSCOUT_USERNAME `
  --insecure `
  --apply
```

화면 계획과 승인표가 같을 때만 `SYNC`를 입력한다. 다르면 다른 값을 입력해 취소한다.

통과 기준:

- CREATE는 POST, ADD는 PUT 수행
- 각 쓰기 후 GET 재조회
- 요청 필드와 서버 필드 일치
- `IsActive=true`
- `results.json`이 `VERIFIED`
- 일부 실패를 성공으로 기록하지 않음

## 17. 단계 12 — 등록 후 Inventory(AFTER) 및 Excel 비교

```powershell
python -m netscout inventory `
  --config $CONFIG `
  --output-dir $OUTPUT `
  --output-file "Inventory_AFTER_REGISTER_$RUN_DATE.xlsx"
```

### 17.1 BEFORE 대 AFTER 비교

비교 Sheet: `Application_List`

기본 키 우선순위:

1. `ApplicationTypeCode`
2. 값이 없거나 불안정하면 `ApplicationShortName`

비교표에 다음 열을 만든다.

| 비교 열 | 계산/판정 |
|---|---|
| Name | ApplicationShortName |
| BEFORE 존재 | BEFORE 키 조회 결과 |
| AFTER 존재 | AFTER 키 조회 결과 |
| Expected Action | plan.json action |
| Actual Result | CREATE/MATCHED/UPDATED/MISSING/UNEXPECTED |
| Type Match | ApplicationType 일치 여부 |
| Parent Match | ParentApplication 일치 여부 |
| IP Match | 주소를 정규화한 집합 비교 |
| Port Match | 포트/범위를 정규화한 집합 비교 |
| Active | IsActive가 true인지 |
| Verdict | 모든 필수 비교가 맞으면 PASS |

Action별 예상:

- CREATE: BEFORE 없음, AFTER 1개 존재, 요청 필드 일치
- MATCHED: BEFORE/AFTER 동일, 불필요한 변경 없음
- ADD_IP: AFTER 주소 = BEFORE 주소 ∪ 요청 신규 주소
- ADD_PORT: AFTER 포트 = BEFORE 포트 ∪ 요청 신규 포트
- CONFLICT: BEFORE와 AFTER 동일, 쓰기 없음

### 17.2 Registration Excel 대 AFTER 비교

등록 승인 Excel을 이름으로 그룹화한 뒤 AFTER Inventory와 비교한다.

- `ApplicationName` ↔ `ApplicationShortName`
- `RegistrationType` ↔ `ApplicationType`
- `ParentApplication` ↔ `ParentApplication`
- SBA `IP` ↔ `ServerAddresses`
- SBA `Port` ↔ `ServerPortRange`
- WKA `Port` ↔ `ApplicationPort`
- CHILD `IP` ↔ `HostAddresses`
- CHILD `Port` ↔ `AdditionalPort`
- URL `URLAddress` ↔ `URLAddress`
- 모든 적용 대상의 `IsActive=true`

IP는 `/32` 표기 차이를 정규화하고, 포트는 순서와 연속 범위 병합 차이를 정규화한 뒤 집합으로 비교한다.

## 18. 단계 13 — 멱등성과 재실행 검증

적용에 사용한 동일 Excel을 `--apply` 없이 다시 실행한다.

```powershell
python .\tools\registration\netscout_sync.py `
  --input ".\output\$RUN_DATE\05_hostname\registration_approved.xlsx" `
  --base-url $BASE_URL `
  --username $env:NETSCOUT_USERNAME `
  --insecure
```

통과 기준:

- 적용된 모든 항목이 `MATCHED`
- applicable change 0건
- POST/PUT/DELETE 없음
- AFTER Inventory를 다시 뽑아도 등록 관련 필드 변화 없음

## 19. 단계 14 — 추가/충돌 시나리오

| ID | 입력 변경 | 예상 결과 |
|---|---|---|
| REG-01 | 신규 이름 + 미사용 IP/Port | CREATE |
| REG-02 | 기존 앱과 완전히 동일 | MATCHED |
| REG-03 | 같은 앱에 신규 IP만 포함 | ADD_IP |
| REG-04 | 같은 앱에 신규 Port만 포함 | ADD_PORT |
| REG-05 | 신규 IP와 Port 모두 포함 | ADD_IP+ADD_PORT |
| REG-06 | 다른 앱이 사용하는 주소 | CONFLICT |
| REG-07 | 같은 이름, 다른 Parent | CONFLICT |
| REG-08 | 같은 이름, 다른 Type | CONFLICT |
| REG-09 | CONFLICT와 CREATE 혼합 + --apply | 전체 적용 전 중단 |
| REG-10 | --apply 후 SYNC 이외 입력 | 취소, 쓰기 0건 |
| REG-11 | 동일 파일 재실행 | MATCHED |
| REG-12 | 기존 IP/Port를 입력에서 빼기 | 자동 삭제/제거하지 않음 |

## 20. 단계 15 — 오류 및 보안 검증

- 잘못된 계정: 인증 오류, 종료 코드 1
- 잘못된 CA: TLS 오류
- 포트 연결 불가: GET 이전 또는 첫 GET에서 중단
- 빈 Inventory: 안전 중단
- 잘못된 IP/Port/이름: 행 번호와 입력 오류
- Formula가 있는 Registration 행: 적용 차단
- 로그/plan/results/XML에 비밀번호 또는 Authorization 값이 없음
- HTTP 오류 시 `last_http_error.txt`에 오류 응답은 남기되 비밀번호는 없음

## 21. 전체 증적 목록

| 단계 | 필수 증적 |
|---|---|
| 환경 | pytest 출력, compileall 결과, TCP 연결 결과 |
| Inventory | PURE, BEFORE, AFTER Excel |
| Top | Top N Excel |
| Targets | 원본/filtered YAML, manual app CSV/YAML |
| Detail | Application별 Detail Excel, Run Summary |
| 후보 | ip_port_candidates.xlsx/yaml |
| Hostname | hostname_lookup.xlsx |
| 등록 입력 | registration_review.xlsx, registration_approved.xlsx |
| 계획 | plan.json |
| 적용 | 요청 XML, results.json |
| 최종 비교 | BEFORE-AFTER 비교표, 승인 Excel-AFTER 비교표 |

## 22. 실행 결과 기록표

| 단계 | 결과 | 산출물 경로 | 건수/요약 | 담당자 | 비고 |
|---|---|---|---|---|---|
| 0 코드/환경 | | | | | |
| 1 PURE Inventory | | | | | |
| 2 Top N | | | | | |
| 3 Detail targets | | | | | |
| 4 Manual app compare/filter | | | | | |
| 5 Detail Batch | | | | | |
| 6 Candidate extraction | | | | | |
| 7 Hostname lookup | | | | | |
| 8 Registration Excel | | | | | |
| 9 Type validation | | | | | |
| 10 BEFORE/Plan | | | | | |
| 11 Apply | | | | | 승인 Gate |
| 12 AFTER/Excel compare | | | | | |
| 13 Idempotency | | | | | |
| 14 Conflict tests | | | | | |
| 15 Error/security | | | | | |

## 23. 최종 릴리즈 판정

다음을 모두 만족해야 v0.89 전체 검증 PASS로 판정한다.

1. 기존 v0.89 자동 테스트와 CLI가 회귀 없이 동작한다.
2. Inventory → Top → Detail → 후보 → Hostname → Registration 변환의 추적성이 유지된다.
3. 각 단계 Excel/YAML의 행 수와 주요 키가 이전 단계와 대조된다.
4. Pure/Current 비교가 수동 등록 Application만 식별한다.
5. 수동 등록 앱 필터가 정확한 Detail 대상만 비활성화한다.
6. Registration 계획의 action이 입력과 현재 서버 상태에 맞다.
7. `--apply` 전에는 쓰기 요청이 없다.
8. 충돌이 하나라도 있으면 전체 적용 전에 중단한다.
9. 승인된 CREATE/ADD 적용 후 GET 재검증과 AFTER Excel 비교가 PASS다.
10. 동일 파일 재실행 결과가 MATCHED이며 추가 쓰기가 없다.
11. 기존 IP/Port 제거와 Application 삭제가 발생하지 않는다.
12. 인증정보가 코드, 로그, Excel, JSON, XML에 노출되지 않는다.

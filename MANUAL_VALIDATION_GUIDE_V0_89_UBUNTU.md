# netscout-auto v0.89 Ubuntu 수동 검증 가이드

Ubuntu 환경에서 v0.89의 조회, 후처리, Hostname, Application 등록 기능을 사람이 직접 검증하는 절차다. 실제 장비 변경은 등록 단계에서 `--apply`를 사용하고 `SYNC`를 입력했을 때만 발생한다.

## 1. 검증 흐름

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
  → Registration 검토 Excel 생성 및 수동 입력
  → Plan-only
  → 실제 등록
  → Inventory AFTER
  → 동일 파일 재실행(MATCHED)
```

## 2. 안전 원칙

1. 테스트 서버에서만 실행한다.
2. 비밀번호를 명령, 문서 또는 YAML에 직접 저장하지 않는다.
3. 등록 도구는 반드시 `--apply` 없이 먼저 실행한다.
4. `CONFLICT`가 있으면 적용하지 않는다.
5. 계획과 승인 Excel이 일치할 때만 `SYNC`를 입력한다.
6. 기존 Application과 겹치지 않는 테스트 IP/Port를 사용한다.
7. 이 도구는 Application 삭제와 기존 IP/Port 제거를 지원하지 않는다.

## 3. Ubuntu 준비

필수 패키지:

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip netcat-openbsd curl
```

Docker가 필요하면 공식 설치 절차에 따라 Docker Engine과 Compose Plugin을 설치한다. 설치 확인:

```bash
docker --version
docker compose version
```

프로젝트로 이동한다.

```bash
cd ~/netscout-api-automation/netscout-auto_v0.89
```

가상환경 사용을 권장한다.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

프로젝트에 `vendor` 의존성이 포함되어 있고 해당 실행 방식을 사용하는 경우에도 가상환경에서 테스트하는 것이 안전하다.

## 4. config 디렉터리 설정

서버와 API 설정을 `config/config.test.yaml`에 직접 입력한다.

```bash
cp ./config/config.example.yaml ./config/config.test.yaml
nano ./config/config.test.yaml
```

테스트 서버 기준 예시:

```yaml
netscout:
  base_url: "https://the-se.iptime.org:8443"
  auth_type: "basic"
  verify_ssl: false
  timeout: 30

api:
  config_applications_endpoint: "/ng1api/ncm/applications"
  devices_endpoint: "/ng1api/ncm/devices"
  device_interfaces_endpoint_template: "/ng1api/ncm/devices/{device_name}/interfaces"
  data_api_endpoint: "/dbonequerydata"
  data_tcp_other_endpoint: "/dbonequerydata"

output:
  directory: "output"
```

조회 CLI가 읽을 계정 파일도 `config` 디렉터리에 만든다.

```bash
nano ./config/.env.local
```

사용자가 직접 입력할 값:

```dotenv
NETSCOUT_USERNAME=administrator
NETSCOUT_PASSWORD=실제비밀번호
```

이 파일은 로컬 테스트 전용이며 Git에 포함하면 안 된다. 이후 모든 조회 명령은 다음 옵션을 직접 사용한다.

```text
--config ./config/config.test.yaml
--env ./config/.env.local
--output-dir ./output
```

등록 CLI는 이 env 파일을 자동으로 읽지 않는다. 등록 명령에는 `--base-url https://the-se.iptime.org:8443`과 `--username administrator`를 직접 쓰고, 비밀번호 프롬프트가 나오면 직접 입력한다.
## 5. 단계 0 — 코드와 연결 검사

```bash
python --version
python -m pytest -q
python -m compileall -q ./V8 ./src ./tools/registration
python -m netscout --help
python ./tools/registration/netscout_sync.py --help
nc -vz the-se.iptime.org 8443
```

HTTPS 수준의 연결도 확인하려면 다음을 사용한다. 자체 서명 인증서를 사용하는 테스트 환경에서만 `-k`를 지정한다.

```bash
curl -k -I --connect-timeout 10 "https://the-se.iptime.org:8443"
```

통과 기준:

- Python 3.11 이상
- pytest 실패 0건
- compileall 오류 없음
- CLI 도움말에 `v0.89` 표시
- `nc` 연결 성공

## 6. 단계 1 — Inventory BASELINE

```bash
python -m netscout inventory \
  --config ./config/config.test.yaml \
  --env ./config/.env.local \
  --output-dir ./output \
  --output-file "Inventory_BASELINE_YYYYMMDD.xlsx"
```

결과:

```text
output/<날짜>/00_inventory/Inventory_BASELINE_<날짜>.xlsx
```

Excel 확인 항목:

- `Application_List`, `Device_List`, `Interface_List`, `Run_Summary` 존재
- Application/Device/Interface 행 수가 로그와 일치
- 사용할 Device Name과 Interface Name/Number 기록
- 등록 테스트 이름이 이미 존재하는지 확인

LibreOffice로 열려면:

```bash
libreoffice "./output/YYYYMMDD/00_inventory/Inventory_BASELINE_YYYYMMDD.xlsx"
```

서버 환경에 GUI가 없으면 파일을 검토 PC로 복사하여 확인한다.

## 7. 단계 2 — Top N

Inventory에서 확인한 실제 Device와 Interface를 사용한다.

```bash
python -m netscout top \
  --config ./config/config.test.yaml \
  --env ./config/.env.local \
  --device "nG1vstream635" \
  --interface "if3" \
  --duration LAST_7_DAYS \
  --top-n 50 \
  --output-dir ./output
```

통과 기준:

- `01_top`에 Excel 생성
- Rank가 1부터 연속
- Bit Rate가 내림차순
- 50행 이하
- Device/Interface가 입력값과 일치

생성 파일 확인:

```bash
find "./output/YYYYMMDD/01_top" -maxdepth 1 -type f -name '*.xlsx' -ls
```

## 8. 단계 3 — Detail 대상 생성

가장 최근 Top N 파일을 변수로 잡는다.

```bash
TOP_XLSX="$(find "./output/YYYYMMDD/01_top" -maxdepth 1 -type f -name '*.xlsx' -printf '%T@ %p\n' | sort -nr | head -1 | cut -d' ' -f2-)"
printf 'TOP_XLSX=%s\n' "$TOP_XLSX"
```

```bash
python -m netscout detail-targets-yaml \
  --top-xlsx "$TOP_XLSX" \
  --detail-top-n 50 \
  --limit 1000
```

결과:

```text
output/<날짜>/02_targets/detail_targets.yaml
```

확인 항목:

- 대상 수 50 이하
- `rank`, `application`, `device`, `interface`, `duration` 존재
- 조회하지 않을 대상은 `enabled: false`

### 단계 3-1 — 수동 등록 Application 화이트리스트 비교 및 제외

앱 등록 전 보관한 순정 Inventory와 현재 BASELINE Inventory를 비교한다. 순정 파일은 `pure_xlsx/pure_inventory.xlsx`에 보관해야 한다.

```bash
python -m netscout compare-inventory-apps \
  --pure-inventory "./pure_xlsx/pure_inventory.xlsx" \
  --current-inventory "./output/YYYYMMDD/00_inventory/Inventory_BASELINE_YYYYMMDD.xlsx" \
  --output "./output/YYYYMMDD/02_targets/manual_registered_apps.csv" \
  --yaml-output "./output/YYYYMMDD/02_targets/manual_registered_apps.yaml" \
  --sheet Application_List \
  --key-column ApplicationTypeCode
```

`PURE`에는 없고 `CURRENT`에만 있는 `ApplicationTypeCode`를 수동 등록 Application으로 판정한다. 생성된 CSV와 YAML을 검토한 뒤 Detail 대상에 반영한다.

```bash
python -m netscout filter-detail-targets \
  --top-yaml "./output/YYYYMMDD/02_targets/detail_targets.yaml" \
  --manual-apps-yaml "./output/YYYYMMDD/02_targets/manual_registered_apps.yaml" \
  --output "./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml" \
  --dry-run

python -m netscout filter-detail-targets \
  --top-yaml "./output/YYYYMMDD/02_targets/detail_targets.yaml" \
  --manual-apps-yaml "./output/YYYYMMDD/02_targets/manual_registered_apps.yaml" \
  --output "./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml"
```

통과 기준:

- 화이트리스트와 일치하는 대상만 `enabled: false`
- 나머지 대상의 순서와 설정은 변경되지 않음
- 이후 Detail 단계는 반드시 `detail_targets_filtered.yaml` 사용

## 9. 단계 4 — Detail Dry-run

```bash
python -m netscout run-detail-targets \
  --config ./config/config.test.yaml \
  --env ./config/.env.local \
  --targets "./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml" \
  --output-dir ./output \
  --default-duration LAST_7_DAYS \
  --dry-run
```

통과 기준:

- 활성 대상이 모두 `planned`
- Device/Interface/Application 값이 예상과 일치
- 아직 Detail API 조회와 Excel 생성은 수행되지 않음

## 10. 단계 5 — Detail 실제 조회

```bash
python -m netscout run-detail-targets \
  --config ./config/config.test.yaml \
  --env ./config/.env.local \
  --targets "./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml" \
  --output-dir ./output \
  --default-duration LAST_7_DAYS
```

결과:

```text
output/<날짜>/03_detail/<Application>/*.xlsx
```

파일 수 확인:

```bash
find "./output/YYYYMMDD/03_detail" -type f -name '*.xlsx' | wc -l
```

통과 기준:

- 모든 활성 Application에 Excel 생성
- 최종 로그 `failed=0`
- Excel의 Application, Device, Interface, 기간이 요청과 일치

SSH 세션 단절이 걱정되면 `tmux` 또는 `screen`에서 실행한다.

## 11. Detail 중단 후 재개

긴 Batch가 중단되면 기존 결과를 삭제하지 않는다. 완료된 Application을 확인한다.

```bash
find "./output/YYYYMMDD/03_detail" -mindepth 2 -type f -name '*.xlsx' -printf '%h\n' | sort -u
```

필터링된 `detail_targets_filtered.yaml`을 복사해 완료된 Application의 `enabled`를 `false`로 바꾼다.

```bash
cp "./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml" \
   "./output/YYYYMMDD/02_targets/detail_targets_resume.yaml"
nano "./output/YYYYMMDD/02_targets/detail_targets_resume.yaml"
```

재개:

```bash
python -m netscout run-detail-targets \
  --config ./config/config.test.yaml \
  --env ./config/.env.local \
  --targets "./output/YYYYMMDD/02_targets/detail_targets_resume.yaml" \
  --output-dir ./output \
  --default-duration LAST_7_DAYS
```

## 12. 단계 6 — IP/Port 후보 추출

HTTPS:

```bash
python -m netscout extract-candidates \
  --detail-root "./output/YYYYMMDD/03_detail/HTTPS" \
  --targets-yaml "./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml" \
  --output "./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.xlsx" \
  --yaml-output "./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.yaml" \
  --applications HTTPS
```

TCP_OTHER:

```bash
python -m netscout extract-candidates \
  --detail-root "./output/YYYYMMDD/03_detail/TCP_OTHER" \
  --targets-yaml "./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml" \
  --output "./output/YYYYMMDD/04_candidates/TCP_OTHER/ip_port_candidates.xlsx" \
  --yaml-output "./output/YYYYMMDD/04_candidates/TCP_OTHER/ip_port_candidates.yaml" \
  --applications TCP_OTHER
```

통과 기준:

- 각 Application의 Excel과 YAML 생성
- IP/Port/Application이 원본 Detail과 일치
- 중복 제거 전후 수량 확인
- `unknown_port_rows` 수동 검토

## 13. 단계 7 — Docker Hostname Lab

실행 설정 생성:

```bash
cp ./config/hostname_lookup.lab.yaml ./config/hostname_lookup.yaml
```

Docker 기동:

```bash
docker compose -f ./docker/hostname-lab/docker-compose.yml up -d --build
docker ps --filter name=netscout-hostname-lab
```

포트 확인:

```bash
curl -k "https://127.0.0.1:5000/search?search=192.0.2.10"
curl -k "https://127.0.0.1:5002/search?search=192.0.2.10"
```

샘플 조회:

```bash
python -m netscout hostname-lookup \
  --input ./docker/hostname-lab/sample_ip_port_candidates.yaml \
  --config ./config/hostname_lookup.yaml
```

HTTPS 후보 전체 조회:

```bash
python -m netscout hostname-lookup \
  --input "./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.yaml" \
  --config ./config/hostname_lookup.yaml
```

결과:

```text
output/<날짜>/05_hostname/HTTPS/hostname_lookup.xlsx
```

통과 기준:

- 컨테이너가 `Up`
- 포트 5000/5002 응답
- 전체 후보 IP 처리
- Lab 데이터의 IP는 Hostname 표시
- 미등록 IP는 빈 Hostname 또는 미매칭 상태

## 14. 단계 8 — Registration 검토 Excel

```bash
python ./tools/registration/make_registration.py \
  --input "./output/YYYYMMDD/05_hostname/HTTPS/hostname_lookup.xlsx" \
  --output "./output/YYYYMMDD/05_hostname/HTTPS/registration_review.xlsx"
```

원본을 보존하고 승인본을 복사한다.

```bash
mkdir -p ./validation_runs
cp "./output/YYYYMMDD/05_hostname/HTTPS/registration_review.xlsx" \
   "./validation_runs/registration_approved_YYYYMMDD.xlsx"
```

LibreOffice가 있는 데스크톱 Ubuntu:

```bash
libreoffice "./validation_runs/registration_approved_YYYYMMDD.xlsx"
```

## 15. Registration Excel 수동 입력

| 열 | 예시 | 확인 사항 |
|---|---|---|
| Register | `Y` | 등록할 행만 Y |
| ApplicationName | `AUTO_V089_MANUAL_01` | 최대 32자, 기존 이름과 중복 금지 |
| RegistrationType | `SBA` | IP/Port 기반 서버 앱 |
| ParentApplication | `TCP` | TCP/UDP/SCTP 또는 기존 앱 |
| Protocol | `TCP` | 부모 프로토콜과 일치 |
| IP | 승인된 테스트 IP | 기존 Application 주소와 중복 금지 |
| Port | 승인된 테스트 Port | 1~65535 |
| Description | 수동 E2E 테스트 | 선택값 |

일반적인 TCP 서버 앱:

```text
Register=Y
RegistrationType=SBA
ParentApplication=TCP
Protocol=TCP
```

다른 Application 이름으로 기존 IP/CIDR을 재사용하면 `CONFLICT`가 발생해야 한다. `Register=N` 행은 등록되지 않는다.

## 16. 단계 9 — 등록 계획만 확인

자체 서명 인증서를 사용하는 테스트 환경 예시:

```bash
python ./tools/registration/netscout_sync.py \
  --input "./validation_runs/registration_approved_YYYYMMDD.xlsx" \
  --base-url "https://the-se.iptime.org:8443" \
  --username "administrator" \
  --insecure
```

사내 CA가 있으면 `--insecure` 대신 다음을 사용한다.

```bash
--ca-file ./internal-ca.pem
```

Action 판정:

| Action | 의미 | 처리 |
|---|---|---|
| CREATE | 신규 Application | 승인 후 적용 가능 |
| MATCHED | 현재 설정과 동일 | 변경 없음 |
| ADD_IP | 기존 앱에 IP 추가 | 승인 후 적용 가능 |
| ADD_PORT | 기존 앱에 Port 추가 | 승인 후 적용 가능 |
| ADD_IP+ADD_PORT | IP와 Port 추가 | 승인 후 적용 가능 |
| CONFLICT | 이름·주소·유형 충돌 | 적용 금지 |

통과 기준:

- 계획의 이름/IP/Port가 Excel과 일치
- 의도하지 않은 대상 없음
- `CONFLICT` 없음
- 장비 쓰기 0건
- `results/sync_<시각>/plan.json` 생성

## 17. 단계 10 — 실제 등록

테스트 장비 변경 승인을 받은 후에만 실행한다.

```bash
python ./tools/registration/netscout_sync.py \
  --input "./validation_runs/registration_approved_YYYYMMDD.xlsx" \
  --base-url "https://the-se.iptime.org:8443" \
  --username "administrator" \
  --insecure \
  --apply
```

계획을 다시 확인하고 정확할 때만 입력한다.

```text
SYNC
```

통과 기준:

- `Applied and verified N change(s)`
- `results.json` 상태가 `VERIFIED`
- XML에 승인한 이름/IP/Port 기록
- 예상과 다르면 `SYNC`를 입력하지 않고 취소

비대화형 파이프 입력도 가능하지만 사람이 계획을 확인하는 검증에서는 권장하지 않는다.

## 18. 단계 11 — Inventory AFTER

```bash
python -m netscout inventory \
  --config ./config/config.test.yaml \
  --env ./config/.env.local \
  --output-dir ./output \
  --output-file "Inventory_AFTER_YYYYMMDD.xlsx"
```

확인 항목:

- 승인한 ApplicationName 존재
- ApplicationType이 `Server-based Apps`
- ParentApplication 일치
- IsActive가 `true`
- Application 수가 CREATE 수만큼 증가

서버 IP/Port의 최종 근거는 적용 XML과 적용 후 GET 검증 결과다. Inventory의 `ApplicationPort`는 실제 서버 Port가 아니라 장비 내부 식별값일 수 있다.

## 19. 단계 12 — MATCHED 재실행

동일 승인 Excel을 `--apply` 없이 다시 실행한다.

```bash
python ./tools/registration/netscout_sync.py \
  --input "./validation_runs/registration_approved_YYYYMMDD.xlsx" \
  --base-url "https://the-se.iptime.org:8443" \
  --username "administrator" \
  --insecure
```

통과 기준:

```text
MATCHED
Plan only; 0 applicable change(s)
```

## 20. 오류 대응

### Permission denied

프로젝트 파일의 소유권과 권한을 확인한다.

```bash
id
ls -la
```

프로젝트 전체에 무조건 `chmod 777`을 적용하지 않는다.

### ModuleNotFoundError

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

### Missing username/password

`config/.env.local`에 사용자가 직접 입력한 계정값이 있는지 확인한다.

```bash
grep -n '^NETSCOUT_USERNAME=' ./config/.env.local
nano ./config/.env.local
```

비밀번호 값 자체를 콘솔에 출력하지 않는다.

### Config API connection failed

```bash
grep -n 'base_url\|verify_ssl' ./config/config.test.yaml
nc -vz the-se.iptime.org 8443
curl -k -I --connect-timeout 10 "https://the-se.iptime.org:8443"
```

### Docker permission denied

현재 사용자가 Docker를 실행할 권한이 있는지 확인한다. 임시로 모든 명령을 `sudo`로 섞어 실행하면 결과 파일 소유권이 달라질 수 있으므로 주의한다.

### CONFLICT

- 실제 적용 금지
- `plan.json`의 충돌 대상 확인
- ApplicationName, IP/CIDR, Port를 BASELINE Inventory와 비교

## 21. 최종 결과 기록표

| 항목 | 결과 | 증적 위치 |
|---|---|---|
| pytest/compile | PASS/FAIL | 콘솔 로그 |
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
| MATCHED | PASS/FAIL/SKIP | `results/.../plan.json` |

## 22. 최종 PASS 기준

1. pytest와 compileall이 통과한다.
2. CLI에 v0.89가 표시된다.
3. Inventory, Top N, Detail 결과가 요청값과 일치한다.
4. 후보 파일이 원본 Detail까지 추적 가능하다.
5. Hostname 성공과 미매칭을 구분할 수 있다.
6. 사람이 승인한 Excel 행만 계획에 포함된다.
7. `CONFLICT`가 있으면 실제 쓰기가 차단된다.
8. `--apply` 전에는 장비 쓰기가 없다.
9. 적용 결과가 `VERIFIED`다.
10. AFTER Inventory에서 대상 앱이 확인된다.
11. 동일 입력 재실행이 `MATCHED`, 변경 0건이다.

## 23. 테스트 종료

Docker Lab 종료:

```bash
docker compose -f ./docker/hostname-lab/docker-compose.yml down
```

계정 정보는 `config/.env.local`에만 저장한다. 검증 후 더 이상 필요하지 않으면 이 파일을 안전하게 제거한다.

가상환경 종료:

```bash
deactivate
```

실제 장비에 생성한 테스트 Application은 이 도구로 자동 삭제하지 않는다. 삭제가 필요하면 NETSCOUT UI 또는 승인된 운영 절차에서 이름과 Type Code를 재확인한 후 처리한다.

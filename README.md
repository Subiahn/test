# netscout-auto v0.89

v0.89는 검증된 조회·분석·후처리 CLI와 Application 등록 동기화 도구를 하나의 정제된 릴리스로 제공합니다. 등록 기능은 안전성을 위해 독립 도구로 유지합니다.

## 핵심 변경점

- 코드·로그·CLI 도움말의 릴리스 버전을 v0.89로 통일
- 실제 계정, 테스트 서버 설정 및 실행 산출물을 배포본에서 제외
- Docker Hostname Lab과 Application 등록 E2E 절차 포함
- 신규 등록 후 GET 검증 및 동일 입력 재실행 `MATCHED` 검증 완료

- `postprocess` 명령 추가
- `extract-candidates` 명령 추가
- `hostname-lookup` 명령 추가
- `postprocess --hostname-lookup`으로 후보 생성 후 Hostname 조회까지 한 번에 실행 가능
- `compare-inventory-apps`로 pure/current Inventory를 비교해 수동 등록 Application 목록 생성
- `filter-detail-targets`로 Detail 대상 YAML에서 이미 등록한 Application을 `enabled: false` 처리
- 기존 `inventory`, `top`, `detail`, `detail-targets-yaml`, `run-detail-targets` 명령은 유지
- `tools/registration/make_registration.py`로 Hostname 조회 결과를 검토용 등록 Excel로 변환
- `tools/registration/netscout_sync.py`로 현재 NETSCOUT 설정과 비교 (`CREATE`, `MATCHED`, `ADD_IP`, `ADD_PORT`, `CONFLICT`)
- 같은 도구에 `--apply`를 지정한 경우에만 POST/PUT을 실행하고 GET으로 재검증
- 삭제와 기존 주소·포트 제거는 자동화하지 않음

등록 기능은 `python -m netscout` 명령이나 공통 API 클라이언트에 합치지 않았습니다. 자세한 사용법은 `tools/registration/README.md`를 참고하십시오.

## 출력 구조

```text
output/
  YYYYMMDD/
    00_inventory/
    01_top/
    02_targets/
    03_detail/
      HTTPS/
      TCP_OTHER/
    04_candidates/
      HTTPS/
        ip_port_candidates.xlsx
        ip_port_candidates.yaml
    05_hostname/
      HTTPS/
        hostname_lookup.xlsx
    run_summary.txt
```

## 준비

Ubuntu:

```bash
cd ~/netscout-auto
cp .env.local.example .env.local
cp config/config.example.yaml config/config.test.yaml
vim .env.local
vim config/config.test.yaml
python3 -m netscout --help
```

Windows PowerShell:

```powershell
cd D:\path\to\netscout-auto
Copy-Item .env.local.example .env.local
Copy-Item config\config.example.yaml config\config.test.yaml
notepad .env.local
notepad config\config.test.yaml
python -m netscout --help
```

## 설정 파일 구분

- `config/config.example.yaml`: 배포용 예제입니다. 실제 서버 주소나 계정 정보는 넣지 않습니다.
- `config/config.test.yaml`: 로컬 테스트 환경에서 공통 `python -m netscout` CLI가 사용하는 파일입니다. 실제 환경값이 들어갈 수 있으므로 Git에 포함하지 않습니다.
- `hostname_lookup.example.yaml`: Hostname 조회 기능의 배포용 예제입니다.
- `hostname_lookup.lab.yaml`: 로컬 Docker/Lab 환경에서 Hostname 조회를 시험할 때 사용합니다.

`tools/registration/netscout_sync.py`는 위 설정 YAML을 읽지 않습니다. 등록 대상 서버는 `--base-url`, 계정은 `--username`과 환경변수 또는 암호 프롬프트, TLS 방식은 `--insecure` 또는 `--ca-file`로 직접 지정합니다.

## 1. Inventory

```bash
python3 -m netscout inventory --config ./config/config.test.yaml --env ./.env.local --output-dir ./output
```

Windows:

```powershell
python -m netscout inventory --config .\config\config.test.yaml --env .\.env.local --output-dir .\output
```

## 2. Top N

```bash
python3 -m netscout top --config ./config/config.test.yaml --env ./.env.local --device "nG1vstream635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir ./output
```

## 3. Detail 대상 YAML 생성

```bash
python3 -m netscout detail-targets-yaml --top-xlsx ./output/YYYYMMDD/01_top/TopN파일.xlsx --detail-top-n 50 --limit 1000
```

기본 저장 위치:

```text
output/YYYYMMDD/02_targets/detail_targets.yaml
```

## 4. Detail Batch 실행

Dry-run:

```bash
python3 -m netscout run-detail-targets --config ./config/config.test.yaml --env ./.env.local --targets ./output/YYYYMMDD/02_targets/detail_targets.yaml --output-dir ./output --default-duration LAST_7_DAYS --dry-run
```

실행:

```bash
python3 -m netscout run-detail-targets --config ./config/config.test.yaml --env ./.env.local --targets ./output/YYYYMMDD/02_targets/detail_targets.yaml --output-dir ./output --default-duration LAST_7_DAYS
```

## 5. 이미 등록한 Application 제외

아무것도 등록하지 않은 순정 NETSCOUT Inventory와 현재 Inventory를 비교해서, 수동 등록된 Application 목록을 만듭니다. 이 단계는 API를 호출하지 않고 Excel만 읽습니다.

```bash
python3 -m netscout compare-inventory-apps --pure-inventory ./pure_xlsx/pure_inventory.xlsx --current-inventory ./output/YYYYMMDD/00_inventory/Netscout_Inventory_YYYYMMDD_HHMMSS.xlsx --output ./output/YYYYMMDD/02_targets/manual_registered_apps.csv --yaml-output ./output/YYYYMMDD/02_targets/manual_registered_apps.yaml
```

그 다음 `detail_targets.yaml`에서 이미 수동 등록된 Application을 `enabled: false`로 바꿔 `detail_targets_filtered.yaml`을 만듭니다.

```bash
python3 -m netscout filter-detail-targets --top-yaml ./output/YYYYMMDD/02_targets/detail_targets.yaml --manual-apps-yaml ./output/YYYYMMDD/02_targets/manual_registered_apps.yaml --output ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml
```

`TCP_OTHER`, `IP_OTHER`, `UDP_OTHER`, `OTHER`, `UNKNOWN`, `IPV6` 계열은 기본적으로 보호되어 자동 비활성화하지 않습니다. 보호를 끄려면 `--disable-protected`를 추가합니다.

이후 Detail Batch와 후보 생성에는 `detail_targets_filtered.yaml`을 사용하는 것을 권장합니다.

```bash
python3 -m netscout run-detail-targets --config ./config/config.test.yaml --env ./.env.local --targets ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml --output-dir ./output --default-duration LAST_7_DAYS
```

## 6. 후보 생성

날짜 폴더 전체를 Application별로 후처리합니다.

```bash
python3 -m netscout postprocess --date-root ./output/YYYYMMDD --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml
```

특정 Application만:

```bash
python3 -m netscout postprocess --date-root ./output/YYYYMMDD --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml --applications HTTPS,TCP_OTHER
```

## 7. 후보 생성 + Hostname 조회

Docker lab 또는 고객사 hostname 조회 서버 설정을 `config/hostname_lookup.yaml`에 둔 뒤 실행합니다.

```bash
python3 -m netscout postprocess --date-root ./output/YYYYMMDD --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml --applications HTTPS --hostname-lookup --hostname-config ./config/hostname_lookup.yaml
```

결과:

```text
output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.xlsx
output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.yaml
output/YYYYMMDD/05_hostname/HTTPS/hostname_lookup.xlsx
```

## 8. Hostname 조회만 단독 실행

```bash
python3 -m netscout hostname-lookup --input ./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.yaml --config ./config/hostname_lookup.yaml
```

## 9. 단일 Detail 폴더 후보 추출

```bash
python3 -m netscout extract-candidates --detail-root ./output/YYYYMMDD/03_detail/HTTPS --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml --output ./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.xlsx --yaml-output ./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.yaml --applications HTTPS
```

## 권장 운영 순서

1. `inventory`
2. `top`
3. `detail-targets-yaml`
4. 필요 시 `compare-inventory-apps`
5. 필요 시 `filter-detail-targets`
6. 사람이 `detail_targets_filtered.yaml` 검토
7. `run-detail-targets`
8. `postprocess`
9. 필요 시 `hostname-lookup` 또는 `postprocess --hostname-lookup`

## 보안

- `.env.local`은 배포/메일/문서에 포함하지 않습니다.
- 비밀번호, token, cookie는 로그와 Excel에 기록하지 않습니다.
- 운영 서버 호출 전에는 반드시 승인 후 실행합니다.

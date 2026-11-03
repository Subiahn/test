# NETSCOUT Automation V8.1 레거시 코어 참고

> v0.89 릴리즈의 기본 실행 진입점은 프로젝트 루트의 `python -m netscout`이다. 이 문서는 보존된 V8.1 코어의 직접 실행 참고 자료다.

V8.1은 V8을 직접 수정하지 않고 복사해서 만든 검증 버전이다.

V8 대비 추가된 내용:

- STEP 2 `Application_BPS_TopN` 시트에 `Short Name` 컬럼 추가
- STEP 2 기존 `BPS` 컬럼을 `Bit Rate` 컬럼으로 정리

현재 코드 기준 `Bit Rate`는 Data API 응답의 `bitRate` 값을 Application 기준으로 합산한 값이다.
Bytes 기반 계산값은 콘솔/API 기준과 혼동될 수 있어 출력하지 않는다.

Data API 문서에서 duration은 `startTime/endTime` 또는 `duration`으로 지정할 수 있고, 기본 timezone은 서버 시간이다.
`LAST_7_DAYS`의 정확한 시작/종료 anchor는 문서에 별도 명시가 없어, 콘솔과 정확히 맞춰 비교할 때는 `--start-time`, `--end-time` 고정 범위를 권장한다.

V8은 현재 조회 전용 기능만 사용한다.

- NETSCOUT에 등록된 Application / Device / Interface 전체 목록 조회
- 특정 Device / Interface 기준 Application Bit Rate Top N 조회
- Top N에서 선택한 Application의 Source / Destination IP / Port 상세 조회
- 여러 Device / Interface 또는 여러 Application을 YAML로 고른 뒤 Batch 실행

신규 후보 분석, Drilldown 분석, raw 분석 명령은 현재 사용하지 않는다.

## 전체 권장 흐름

처음 사용하는 경우 아래 순서로 진행한다.

1. `--inventory`로 현재 NETSCOUT Application / Device / Interface 목록을 확인한다.
2. `--targets-yaml`로 `config\targets.yaml` 실행 대상 YAML을 만든다.
3. 사람이 YAML을 열어 필요 없는 Interface를 `enabled: false`로 제외한다.
4. `--run-targets --dry-run`으로 실행 예정 대상을 확인한다.
5. `--run-targets`로 대상별 Application Top N Excel을 만든다.
6. Top N Excel을 보고 상세 확인할 Application을 고른다.
7. `--detail-targets-yaml`로 `config\detail_targets.yaml` Application Detail 대상 YAML을 만든다.
8. 사람이 YAML을 열어 상세 조회할 Application만 남긴다.
9. `--run-detail-targets --dry-run`으로 실행 예정 대상을 확인한다.
10. `--run-detail-targets`로 Application별 IP / Port 상세 Excel을 만든다.

V8은 Application 생성/수정/삭제를 하지 않는다. 모든 기능은 조회와 Excel/YAML 생성이다.

## 0. 실행 전 준비

V8.1은 폴더 이름에 점(`.`)이 들어가 있으므로 `python -m V8.1...` 형태로 실행하지 않는다.

권장 실행 방식은 아래 둘 중 하나다.

1. `V8.1` 폴더로 이동한 뒤 `python .\v8_cli.py ...` 실행
2. 프로젝트 루트에서 `python V8.1\v8_cli.py ...` 실행

### Windows PowerShell 준비

```powershell
cd D:\yury\netscout-api-automation
```

필요 파일:

- `config\config.test.yaml`
- `.env.local`

`.env.local` 예시:

```text
NETSCOUT_USERNAME=dummy-user
NETSCOUT_PASSWORD=CHANGE_ME
```

실제 비밀번호는 코드, README, 로그, Excel에 기록하지 않는다.

### Windows PowerShell 빠른 실행 예

아래 예시는 `V8.1` 폴더 안에서 실행하는 기준이다.

```powershell
cd D:\yury\netscout-api-automation\V8.1
```

Inventory:

```powershell
python .\v8_cli.py inventory --config ..\config\config.test.yaml --env ..\.env.local --output-dir ..\output\v81
```

Application Top 50, duration 방식:

```powershell
python .\v8_cli.py top --config ..\config\config.test.yaml --env ..\.env.local --device "nG1vstream635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir ..\output\v81
```

Application Top 50, 콘솔 비교용 고정 시간 방식:

```powershell
python .\v8_cli.py top --config ..\config\config.test.yaml --env ..\.env.local --device "nG1vstream635" --interface "if3" --start-time 2026-09-23_00:00:00 --end-time 2026-09-30_00:00:00 --resolution NO_RESOLUTION --top-n 50 --output-dir ..\output\v81
```

특정 Application 상세 IP / Port 분석:

```powershell
python .\v8_cli.py detail --config ..\config\config.test.yaml --env ..\.env.local --device "nG1vstream635" --interface "if3" --application "HTTPS" --start-time 2026-09-23_00:00:00 --end-time 2026-09-30_00:00:00 --resolution NO_RESOLUTION --limit 1000 --output-dir ..\output\v81
```

`12 AM to 12 AM` 기준이면 시작/종료 시간을 모두 `00:00:00`으로 입력한다.

```text
12 AM = 00:00:00
12 PM = 12:00:00
```

### Ubuntu 준비

Ubuntu에서는 Windows 경로 구분자 `\`를 쓰지 말고 `/`를 사용한다.

프로젝트 루트에서 실행하는 예:

```bash
cd ~/netscout-api-automation
```

또는 배포 폴더 이름이 `netscout-auto`라면:

```bash
cd ~/netscout-auto
```

Inventory:

```bash
python3 V8.1/v8_cli.py inventory --config ./config/config.test.yaml --env ./.env.local --output-dir ./output/v81
```

Application Top 50, duration 방식:

```bash
python3 V8.1/v8_cli.py top --config ./config/config.test.yaml --env ./.env.local --device "nG1vstream635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir ./output/v81
```

Application Top 50, 콘솔 비교용 고정 시간 방식:

```bash
python3 V8.1/v8_cli.py top --config ./config/config.test.yaml --env ./.env.local --device "nG1vstream635" --interface "if3" --start-time 2026-09-23_00:00:00 --end-time 2026-09-30_00:00:00 --resolution NO_RESOLUTION --top-n 50 --output-dir ./output/v81
```

특정 Application 상세 IP / Port 분석:

```bash
python3 V8.1/v8_cli.py detail --config ./config/config.test.yaml --env ./.env.local --device "nG1vstream635" --interface "if3" --application "HTTPS" --start-time 2026-09-23_00:00:00 --end-time 2026-09-30_00:00:00 --resolution NO_RESOLUTION --limit 1000 --output-dir ./output/v81
```

Ubuntu에서 `No module named ...` 오류가 나면 현재 위치가 프로젝트 루트인지 확인한다.

```bash
pwd
ls
ls V8.1/v8_cli.py
```

## 1. Help 확인

전체 명령어 확인:

```powershell
python .\v8_cli.py --help
```

현재 사용하는 주요 명령:

- `--inventory` 또는 `inventory`
- `--top` 또는 `top`
- `--detail` 또는 `detail`
- `--targets-yaml` 또는 `targets-yaml`
- `--run-targets` 또는 `run-targets`
- `--detail-targets-yaml` 또는 `detail-targets-yaml`
- `--run-detail-targets` 또는 `run-detail-targets`

플래그형 명령과 서브커맨드형 명령은 같은 기능이다.

예를 들어 아래 두 명령은 동일하다.

```powershell
python .\v8_cli.py --top --config ..\config\config.test.yaml --env ..\.env.local --device "nG1vstream635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir ..\output\v81
python .\v8_cli.py top --config ..\config\config.test.yaml --env ..\.env.local --device "nG1vstream635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir ..\output\v81
```

## 2. Inventory 조회

NETSCOUT에 등록된 Application, Device, Interface 목록을 Excel로 저장한다.

### 실행 명령

```powershell
python -m V8.v8_cli --inventory --config config\config.test.yaml --env .env.local --output-dir output\v8
```

동일한 명령을 서브커맨드 형태로도 실행할 수 있다.

```powershell
python -m V8.v8_cli inventory --config config\config.test.yaml --env .env.local --output-dir output\v8
```

### 옵션 설명

- `--config`: NETSCOUT base_url, endpoint, timeout, verify_ssl 설정 파일
- `--env`: 계정 정보를 읽을 env 파일
- `--output-dir`: Excel 저장 폴더

### 출력 파일 예

```text
output\v8\Netscout_V8_Inventory_YYYYMMDD_HHMMSS.xlsx
```

### 출력 시트

- `Application_List`
- `Device_List`
- `Interface_List`
- `Run_Summary`

`Application_List`에는 아래 Application 메타데이터가 포함된다.

- `ApplicationShortName`
- `ApplicationLongName`
- `ApplicationType`
- `ApplicationTypeCode`
- `ParentApplication`
- `ApplicationPort`
- `ApplicationGroup`
- `IsActive`

## 3. Application Top N 조회

특정 Device / Interface 기준으로 Application 트래픽을 Bit Rate 기준 내림차순 정렬해서 Top N으로 저장한다.

### 실행 명령

```powershell
python -m V8.v8_cli --top --config config\config.test.yaml --env .env.local --device "vStream_635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir output\v8
```

서브커맨드 형태:

```powershell
python -m V8.v8_cli top --config config\config.test.yaml --env .env.local --device "vStream_635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir output\v8
```

### 옵션 설명

- `--device`: 조회할 Device Name
- `--interface`: Interface Name 또는 Interface Number. 예: `if3`, `3`
- `--duration`: NETSCOUT duration. 예: `LAST_HOUR`, `LAST_24_HOURS`, `LAST_7_DAYS`
- `--top-n`: 출력할 Application 개수. 예: `50`
- `--output-dir`: Excel 저장 폴더

### 절대 시간 조회

duration 대신 특정 시작/종료 시간을 지정할 수 있다.

```powershell
python -m V8.v8_cli --top --config config\config.test.yaml --env .env.local --device "vStream_635" --interface "if3" --start-time 2026-08-24_12:00:00 --end-time 2026-08-31_12:00:00 --resolution NO_RESOLUTION --top-n 50 --output-dir output\v8
```

`--duration`과 `--start-time/--end-time`은 동시에 사용하지 않는다.

### 출력 파일 예

```text
output\v8\vStream_635_if3_Top50_LAST_7_DAYS_YYYYMMDD_HHMMSS.xlsx
```

### 출력 시트

- `Application_BPS_TopN`

### 주요 컬럼

- `Rank`
- `Application Name`
- `Application Type Code`
- `Application Type`
- `Parent Application`
- `Application Group`
- `App ID`
- `Bit Rate`
- `Bytes`
- `Packets`
- `Device Name`
- `Device IP`
- `Interface Name`
- `Interface Alias`
- `Interface Number`

`Siemens_S7` 같은 Application이 나오는 것은 해당 Interface에서 실제로 관측된 트래픽이기 때문이다. 미등록 후보라는 뜻은 아니며, Config API Application List와 매칭되는 등록 Application일 수 있다.

## 4. Batch 대상 YAML 생성

여러 Device / Interface를 한 번에 돌리기 전에, 먼저 Inventory 기준으로 대상 YAML 파일을 만들 수 있다.

이 단계는 Top / Detail API를 호출하지 않는다. Config API로 Device / Interface 목록만 읽고 YAML 파일을 생성한다.

기본 저장 위치는 `config\targets.yaml`이다. 이 파일은 사람이 편집하는 실행 설정에 가깝기 때문에 `output`이 아니라 `config` 아래에 둔다.

사람이 YAML을 열어서 필요 없는 대상은 `enabled: false`로 바꾸거나 줄을 삭제한 뒤, 나중에 Batch 실행에 사용할 수 있다.

Device / Interface가 여러 개 나오면 `targets` 아래에 여러 항목으로 생성된다.

```yaml
targets:
- enabled: true
  device: vStream_635
  interface: if3
  interface_number: 3
- enabled: true
  device: vStream_635
  interface: if4
  interface_number: 4
- enabled: true
  device: vStream_636
  interface: if3
  interface_number: 3
```

사람이 확인 후 제외할 대상은 아래처럼 바꾼다.

```yaml
- enabled: false
  device: vStream_635
  interface: if4
  memo: 이번 실행에서 제외
```

### 전체 대상 YAML 생성

```powershell
python -m V8.v8_cli --targets-yaml --config config\config.test.yaml --env .env.local --duration LAST_7_DAYS --top-n 50
```

서브커맨드 형태:

```powershell
python -m V8.v8_cli targets-yaml --config config\config.test.yaml --env .env.local --duration LAST_7_DAYS --top-n 50
```

위 명령은 기본적으로 아래 파일을 만든다.

```text
config\targets.yaml
```

다른 위치에 만들고 싶으면 `--output`을 직접 지정한다.

```powershell
python -m V8.v8_cli --targets-yaml --config config\config.test.yaml --env .env.local --output config\my_targets.yaml --duration LAST_7_DAYS --top-n 50
```

### 특정 Device만 YAML 생성

```powershell
python -m V8.v8_cli --targets-yaml --config config\config.test.yaml --env .env.local --device-filter "vStream_635" --duration LAST_7_DAYS --top-n 50
```

### 특정 Interface만 YAML 생성

```powershell
python -m V8.v8_cli --targets-yaml --config config\config.test.yaml --env .env.local --device-filter "vStream_635" --interface-filter "if3" --duration LAST_7_DAYS --top-n 50
```

### 생성되는 YAML 예시

```yaml
version: 1
generated_at: '2026-09-21 13:30:00'
description: Edit enabled/include_applications/exclude_applications before running batch.
defaults:
  duration: LAST_7_DAYS
  top_n: 50
  run_detail: false
  detail_top_n: 10
targets:
- enabled: true
  device: vStream_635
  device_ip: 192.0.2.10
  interface: if3
  interface_alias: market
  interface_number: 3
  ifn: 3
  status: Active
  duration: LAST_7_DAYS
  top_n: 50
  run_detail: false
  detail_top_n: 10
  include_applications: []
  exclude_applications: []
  memo: ''
```

편집 기준:

- 실행 제외: `enabled: false`
- 특정 Application만 상세 조회할 예정이면 `include_applications`에 이름 추가
- 제외할 Application은 `exclude_applications`에 이름 추가
- 사람이 판단한 메모는 `memo`에 작성

## 5. YAML을 읽어서 Top N 자동 실행

사람이 `targets.yaml`을 확인/수정한 뒤, `enabled: true`인 대상만 자동으로 Top N을 실행할 수 있다.

먼저 실제 호출 없이 실행 예정 대상만 확인한다.

```powershell
python -m V8.v8_cli --run-targets --config config\config.test.yaml --env .env.local --targets config\targets.yaml --output-dir output\v8 --dry-run
```

dry-run 로그에서 대상이 맞으면 실제 실행한다.

```powershell
python -m V8.v8_cli --run-targets --config config\config.test.yaml --env .env.local --targets config\targets.yaml --output-dir output\v8
```

서브커맨드 형태:

```powershell
python -m V8.v8_cli run-targets --config config\config.test.yaml --env .env.local --targets config\targets.yaml --output-dir output\v8
```

동작 방식:

1. `targets.yaml`을 읽는다.
2. `enabled: false` 대상은 건너뛴다.
3. 각 enabled 대상의 `device`, `interface`, `duration`, `top_n` 값을 사용한다.
4. 대상별로 `Application_BPS_TopN` Excel을 생성한다.
5. 실패한 대상은 로그에 남기고 다음 대상으로 계속 진행한다.

현재 `run-targets`는 Top N 자동 실행만 수행한다. Top N 결과에서 Detail 대상을 고르는 작업은 아래 `detail-targets-yaml` 흐름을 사용한다.

## 6. Top N 결과로 Detail 대상 YAML 생성

Top N Excel의 `Application_BPS_TopN` 시트를 읽어서 Detail 실행 대상 YAML을 만든다.

이 YAML에는 Top 1부터 Top 50까지 Rank 순서가 유지되며, 사람이 참고할 수 있도록 `bps`, `bytes`, `packets`, Application Type 정보도 같이 들어간다.

이 단계는 API를 호출하지 않는다. 이미 생성된 Top N Excel만 읽는다.

기본 저장 위치는 `config\detail_targets.yaml`이다.

중요한 해석:

- `rank`: STEP 2 Top N 순위
- `bps`: STEP 2에서 본 Application 기준 bitRate 합산값
- `bytes`, `packets`: STEP 2에서 본 Application 기준 사용량
- 위 값들은 사람이 어떤 Application을 상세 조회할지 고르는 참고값이다.
- 실제 Detail 조회 필터는 `device`, `interface`, `application`, `duration`을 기준으로 한다.

### 실행 명령

먼저 `output\v8` 폴더에서 실제 생성된 Top N Excel 파일명을 확인한다.

```powershell
Get-ChildItem output\v8\*Top50*.xlsx | Sort-Object LastWriteTime -Descending
```

아래 명령의 `vStream_635_if3_Top50_LAST_7_DAYS_YYYYMMDD_HHMMSS.xlsx`는 예시 파일명이다.
`YYYYMMDD_HHMMSS`를 그대로 쓰지 말고, 위에서 확인한 실제 파일명으로 바꿔서 실행한다.

```powershell
python -m V8.v8_cli --detail-targets-yaml --top-xlsx output\v8\vStream_635_if3_Top50_LAST_7_DAYS_YYYYMMDD_HHMMSS.xlsx --detail-top-n 50 --limit 1000
```

서브커맨드 형태:

```powershell
python -m V8.v8_cli detail-targets-yaml --top-xlsx output\v8\vStream_635_if3_Top50_LAST_7_DAYS_YYYYMMDD_HHMMSS.xlsx --detail-top-n 50 --limit 1000
```

Detail Batch에서 Source IP / Destination IP 기준 중복 제거를 기본으로 켜고 싶으면 `--dedup-ip`를 붙여서 YAML을 생성한다.

```powershell
python -m V8.v8_cli --detail-targets-yaml --top-xlsx output\v8\vStream_635_if3_Top50_LAST_7_DAYS_YYYYMMDD_HHMMSS.xlsx --detail-top-n 50 --limit 1000 --dedup-ip
```

위 명령은 기본적으로 아래 파일을 만든다.

```text
config\detail_targets.yaml
```

### 옵션 설명

- `--top-xlsx`: STEP 2 또는 `run-targets`로 생성한 Top N Excel 파일
- `--output`: 생성할 Detail targets YAML 파일
- `--detail-top-n`: Rank 기준으로 YAML에 넣을 Application 개수. 기본값 `50`
- `--limit`: 각 Application Detail 조회 시 저장할 상세 row 제한. 기본값 `1000`
- `--dedup-ip`: Detail 실행 시 Source IP / Destination IP 조합 중복 제거를 기본값으로 넣는다.

### 생성되는 YAML 예시

```yaml
version: 1
generated_at: '2026-09-21 14:00:00'
source_top_xlsx: output\v8\vStream_635_if3_Top50_LAST_7_DAYS_YYYYMMDD_HHMMSS.xlsx
description: Rank/Bit Rate are copied from Application_BPS_TopN for human review before detail batch.
defaults:
  limit: 1000
  dedup_ip: false
targets:
- enabled: true
  rank: 1
  application: HTTPS
  app_id: 50462829
  bps: 2930306.8441403
  bytes: 123456789
  packets: 12345
  application_type: Well Known Apps
  application_type_code: HTTPS
  parent_application: TCP
  application_group: Web
  device: vStream_635
  device_ip: 192.0.2.10
  interface: if3
  interface_alias: market
  interface_number: 3
  duration: ''
  limit: 1000
  dedup_ip: false
  memo: ''
```

편집 기준:

- Detail 조회 제외: `enabled: false`
- 조회 기간을 Application별로 다르게 줄 때: `duration: LAST_HOUR`처럼 입력
- 전체 기본 기간을 쓰려면 `duration: ''` 그대로 둔다.
- `bps`는 Top N에서 가져온 참고값이다. Detail API 필터에는 쓰지 않는다.
- `limit`은 상세 row를 bitRate 내림차순으로 몇 개까지 저장할지 정한다.
- `dedup_ip: true`이면 Source IP / Destination IP 조합이 같은 row 중 bitRate가 가장 높은 row만 남긴다.
- 중복 제거 기준에 Port는 포함하지 않는다.
- `A -> B`와 `B -> A`는 서로 다른 방향의 통신으로 보고 둘 다 유지한다.

## 7. Detail targets YAML을 읽어서 상세 자동 실행

사람이 `detail_targets.yaml`을 확인/수정한 뒤, `enabled: true`인 Application만 Detail 조회한다.

먼저 실제 호출 없이 실행 예정 대상만 확인한다.

```powershell
python -m V8.v8_cli --run-detail-targets --config config\config.test.yaml --env .env.local --targets config\detail_targets.yaml --output-dir output\v8 --default-duration LAST_7_DAYS --dry-run
```

dry-run 결과가 맞으면 실제 실행한다.

```powershell
python -m V8.v8_cli --run-detail-targets --config config\config.test.yaml --env .env.local --targets config\detail_targets.yaml --output-dir output\v8 --default-duration LAST_7_DAYS
```

서브커맨드 형태:

```powershell
python -m V8.v8_cli run-detail-targets --config config\config.test.yaml --env .env.local --targets config\detail_targets.yaml --output-dir output\v8 --default-duration LAST_7_DAYS
```

Batch Detail 결과도 target의 Application 이름별 폴더에 저장된다.
예를 들어 YAML에 `application: HTTPS`와 `application: TCP_OTHER`가 있으면 각각 `output\v8\HTTPS\`, `output\v8\TCP_OTHER\` 아래에 Excel이 생성된다.

### YAML에서 중복 제거 적용하기

`config\detail_targets.yaml` 안에서 전체 기본값으로 중복 제거를 켤 수 있다.

```yaml
defaults:
  limit: 1000
  dedup_ip: true
```

특정 Application만 중복 제거를 켜거나 끄고 싶으면 target별 `dedup_ip` 값을 수정한다.
target별 값이 `defaults.dedup_ip`보다 우선한다.

```yaml
targets:
- enabled: true
  rank: 1
  application: HTTPS
  device: vStream_635
  interface: if3
  duration: ''
  limit: 1000
  dedup_ip: true

- enabled: true
  rank: 2
  application: TCP_OTHER
  device: vStream_635
  interface: if3
  duration: ''
  limit: 1000
  dedup_ip: false
```

중복 제거가 켜진 target도 원본 Detail은 그대로 유지한다.
그리고 IP 확인용 요약 시트가 추가로 생성된다.

- `Application_Detail`: 원본 상세 row
- `Application_Detail_After_Dedup`
- `Unique_IP_Pairs`: Source IP / Destination IP 기준 요약
- `Unique_IP_Port_Pairs`: Source IP / Source Port / Destination IP / Destination Port 기준 요약

중복 제거가 꺼진 target은 아래 시트 하나만 생성된다.

- `Application_Detail`

동작 방식:

1. `detail_targets.yaml`을 읽는다.
2. `enabled: false` 대상은 건너뛴다.
3. 각 enabled 대상의 `device`, `interface`, `application`, `duration`, `limit`, `dedup_ip` 값을 사용한다.
4. 대상별로 Application Detail Excel을 생성한다.
5. 실패한 대상은 로그에 남기고 다음 대상으로 계속 진행한다.

실패 처리는 Batch 운영을 위해 느슨하게 동작한다.
예를 들어 50개 Application 중 1개가 HTTP 400/500으로 실패해도 나머지 Application 조회는 계속 진행한다.

## 8. 선택 Application 상세 조회

Top N 결과에서 Application 하나를 선택한 뒤, 해당 Application의 Source / Destination IP와 Port 상세를 조회한다.

### HTTPS 상세 조회 예시

```powershell
python -m V8.v8_cli --detail --config config\config.test.yaml --env .env.local --device "vStream_635" --interface "if3" --application "HTTPS" --duration LAST_7_DAYS --output-dir output\v8
```

서브커맨드 형태:

```powershell
python -m V8.v8_cli detail --config config\config.test.yaml --env .env.local --device "vStream_635" --interface "if3" --application "HTTPS" --duration LAST_7_DAYS --output-dir output\v8
```

### 옵션 설명

- `--device`: 조회할 Device Name
- `--interface`: Interface Name 또는 Interface Number
- `--application`: 상세 조회할 Application 이름. 예: `HTTPS`, `BING`, `TCP_OTHER`
- `--duration`: NETSCOUT duration
- `--limit`: bitRate 내림차순 기준 상세 row 제한. 기본값 `1000`
- `--dedup-ip`: Source IP / Destination IP 조합 기준으로 중복 제거
- `--output-dir`: Excel 저장 폴더

### 중복 제거 예시

```powershell
python -m V8.v8_cli --detail --config config\config.test.yaml --env .env.local --device "vStream_635" --interface "if3" --application "HTTPS" --duration LAST_7_DAYS --dedup-ip --output-dir output\v8
```

### 상세 row 개수 제한 예시

```powershell
python -m V8.v8_cli --detail --config config\config.test.yaml --env .env.local --device "vStream_635" --interface "if3" --application "HTTPS" --duration LAST_7_DAYS --limit 500 --output-dir output\v8
```

### 절대 시간 상세 조회

```powershell
python -m V8.v8_cli --detail --config config\config.test.yaml --env .env.local --device "vStream_635" --interface "if3" --application "HTTPS" --start-time 2026-08-24_12:00:00 --end-time 2026-08-31_12:00:00 --resolution NO_RESOLUTION --output-dir output\v8
```

### 출력 파일 예

```text
output\v8\HTTPS\vStream_635_if3_HTTPS_Detail_LAST_7_DAYS_YYYYMMDD_HHMMSS.xlsx
```

Detail 결과는 Application별 폴더 아래에 저장된다.
예를 들어 `--application "HTTPS"`는 `output\v8\HTTPS\` 아래에 저장되고,
`--application "TCP_OTHER"`는 `output\v8\TCP_OTHER\` 아래에 저장된다.

### 주요 컬럼

- `Application Name`
- `App ID`
- `Source IP`
- `Source Port`
- `Destination IP`
- `Destination Port`
- `bitRate`
- `bitRateIn`
- `bitRateOut`
- `Bytes`
- `Packets`
- `Device Name`
- `Device IP`
- `Interface Name`
- `Interface Alias`
- `Interface Number`

## 9. API 흐름

Inventory:

- `GET /ng1api/ncm/applications`
- `GET /ng1api/ncm/devices`
- `GET /ng1api/ncm/devices/{deviceName}/interfaces`

Top N / Detail:

- `POST /dbonequerydata`
- `NetworkParameter=APPLICATION`
- `SelectColumnList`에는 `appId`, `appId_String`, `srcAddress`, `srcPort`, `destAddress`, `destPort`, `bitRate`, `bitRateIn`, `bitRateOut` 등이 포함된다.

## 10. 자주 헷갈리는 점

### Application Type은 어디서 오나?

Top N의 `Application Name`은 사람이 보기 쉽도록 Config API Application List의 Short Name(`application_short_name`)을 우선 표시한다.
Short Name이 없으면 Data API의 `appId_String` 값을 fallback으로 표시한다.

Top N의 `Application Type Code`, `Application Type`, `Parent Application`, `Application Group`은 Config API Application List와 `appId_String`을 매칭해서 붙인다.
상세 조회 Batch YAML을 만들 때는 API Query 안정성을 위해 `Application Type Code`를 우선 사용하고, 없을 때만 `Application Name`을 사용한다.

### Bit Rate는 무엇인가?

Excel의 `Bit Rate`는 Data API 응답의 `bitRate` 값을 Application 기준으로 합산한 값이다.

### `Source Port`가 0으로 나올 수 있나?

나올 수 있다. 이 경우 `Destination Port`와 함께 봐야 한다.

### `Monitor IP / Interface`와 `Source / Destination IP`는 다른가?

다르다.

- Device / Interface는 조회 기준이 되는 관측 지점이다.
- Source IP / Destination IP는 해당 관측 지점에서 보인 실제 통신 endpoint다.

## 11. 주의사항

- 비밀번호, token, cookie는 코드/문서/로그에 기록하지 않는다.
- `.env.local`은 Git에 포함하지 않는다.
- `config.test.yaml`에는 비밀번호를 넣지 않는다.
- 운영 서버 호출은 별도 승인 후 수행한다.
- V8은 조회용이며 Application 생성/수정/삭제 API는 호출하지 않는다.

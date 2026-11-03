# netscout-auto tools 사용 가이드

이 폴더의 도구는 API 호출 없이 이미 생성된 Excel/YAML을 후처리한다.

포함 도구:

- `registration/`: 기존 V8 CLI와 분리된 Application 등록 Excel 생성, 계획 및 선택적 적용 도구
- `nca/`: 향후 두 번째 Hostname 호출 모듈을 위한 예약 디렉터리
- `compare_inventory_apps.py`: 순정/현재 Inventory Excel을 비교해 수동 등록 Application 후보를 만든다.
- `filter_top_yaml_by_manual_apps.py`: Top 50 Detail 대상 YAML에서 수동 등록 Application을 `enabled: false`로 표시한다.
- `extract_ip_port_candidates.py`: Application별 Detail Excel에서 IP/Port 후보와 중복 제거 IP 목록을 만든다.
- `lookup_hostnames.py`: 중복 제거 IP 목록을 hostname 조회 사이트에 일괄 조회한다.

## Inventory Application 비교 도구

`compare_inventory_apps.py`는 API를 호출하지 않고 Inventory Excel 2개만 비교한다.

목적:

1. 아무것도 수동 등록하지 않은 순정 NETSCOUT에서 Inventory를 뽑는다.
2. Application을 수동 등록한 뒤 현재 Inventory를 다시 뽑는다.
3. 두 파일의 `Application_List` 시트를 `ApplicationTypeCode` 기준으로 비교한다.
4. 순정 파일에 없고 현재 파일에만 있는 Application을 CSV로 저장한다.

## Ubuntu

`netscout-auto` 폴더에서 실행한다.

```bash
python3 tools/compare_inventory_apps.py \
  --pure-inventory ./output/pure_inventory.xlsx \
  --current-inventory ./output/current_inventory.xlsx \
  --output ./output/manual_registered_apps.csv \
  --yaml-output ./config/manual_registered_apps.yaml
```

## Windows PowerShell

```powershell
python .\tools\compare_inventory_apps.py `
  --pure-inventory .\output\pure_inventory.xlsx `
  --current-inventory .\output\current_inventory.xlsx `
  --output .\output\manual_registered_apps.csv `
  --yaml-output .\config\manual_registered_apps.yaml
```

## 기본 비교 기준

- Sheet: `Application_List`
- Key column: `ApplicationTypeCode`

다른 컬럼으로 비교해야 할 때:

```bash
python3 tools/compare_inventory_apps.py \
  --pure-inventory ./output/pure_inventory.xlsx \
  --current-inventory ./output/current_inventory.xlsx \
  --output ./output/manual_registered_apps.csv \
  --key-column ApplicationShortName
```

## 결과 CSV

기본 컬럼:

- `ApplicationShortName`
- `ApplicationLongName`
- `ApplicationType`
- `ApplicationTypeCode`
- `ParentApplication`
- `ApplicationPort`
- `AdditionalPort`
- `ApplicationGroup`
- `IsActive`
- `Logged`
- `ResponseTime`
- `Session`
- `RecordingType`

이 CSV는 다음 단계에서 Top N 결과에서 수동 등록 Application을 제외하는 whitelist/diff 파일로 사용할 수 있다.

## 결과 YAML

`--yaml-output`을 지정하면 사람이 편집하기 쉬운 YAML도 같이 생성한다.

예:

```yaml
description: Pure Inventory에는 없고 Current Inventory에만 있는 수동 등록 Application 후보.
compare_key: ApplicationTypeCode
manual_registered_applications:
  - enabled: true
    key: MY_APP
    application_short_name: MY_APP
    application_type_code: MY_APP
    application_port: 12345
```

`enabled` 값을 사람이 `false`로 바꾸면, 다음 단계에서 Top N 제외 목록으로 사용할 때 일부 Application만 선택적으로 제외할 수 있다.

## Top 50 YAML에서 수동 등록 Application 비활성화

`detail-targets-yaml`로 만든 Top 50 YAML과 `manual_registered_apps.yaml`을 비교해서,
수동 등록 Application과 매칭되는 target을 `enabled: false`로 바꿀 수 있다.

Ubuntu:

```bash
python3 tools/filter_top_yaml_by_manual_apps.py \
  --top-yaml ./config/detail_targets.yaml \
  --manual-apps-yaml ./config/manual_registered_apps.yaml \
  --output ./config/detail_targets_filtered.yaml
```

Windows PowerShell:

```powershell
python .\tools\filter_top_yaml_by_manual_apps.py `
  --top-yaml .\config\detail_targets.yaml `
  --manual-apps-yaml .\config\manual_registered_apps.yaml `
  --output .\config\detail_targets_filtered.yaml
```

기본 동작:

- `application_type_code` / `key` / `application` 후보를 비교한다.
- 매칭되면 `enabled: false`를 넣는다.
- `TCP_OTHER`, `UDP_OTHER`, `IP_OTHER`, `OTHER`, `UNKNOWN`, `IPV6` 계열은 기본적으로 보호한다.

보호 대상까지 비활성화하려면 `--disable-protected`를 추가한다.

먼저 통계만 보고 싶으면 `--dry-run`을 사용한다.

## Detail Excel에서 IP/Port 후보 추출

`run-detail-targets`로 생성된 Application별 Detail Excel을 읽어서 hostname 조회용 IP/Port 후보를 만든다.
Detail 조회와 후보 추출은 분리해서 실행한다.

v0.85 Detail 결과는 보통 날짜 폴더 아래에 저장된다.

```text
output/YYYYMMDD/HTTPS/...
output/YYYYMMDD/TCP_OTHER/...
```

이 도구는 `--detail-root` 아래를 재귀적으로 읽는다.
운영 기준은 Application별 폴더를 지정해서 후보 파일도 같은 Application 폴더 아래에 저장하는 것이다.

Ubuntu:

```bash
python3 tools/extract_ip_port_candidates.py \
  --detail-root ./output/20261007/HTTPS \
  --targets-yaml ./config/detail_targets_filtered.yaml \
  --output ./output/20261007/HTTPS/ip_port_candidates.xlsx \
  --yaml-output ./output/20261007/HTTPS/ip_port_candidates.yaml \
  --applications HTTPS
```

Windows PowerShell:

```powershell
python .\tools\extract_ip_port_candidates.py `
  --detail-root .\output\20261007\HTTPS `
  --targets-yaml .\config\detail_targets_filtered.yaml `
  --output .\output\20261007\HTTPS\ip_port_candidates.xlsx `
  --yaml-output .\output\20261007\HTTPS\ip_port_candidates.yaml `
  --applications HTTPS
```

출력 Excel sheet:

- `IP_Port_Candidates`: Application / Direction / IP / Port 기준 집계
- `Unknown_Port_Rows`: Source Port와 Destination Port가 둘 다 0이거나 확인이 필요한 원본 row
- `Unique_IPs`: hostname 조회용 중복 제거 IP 목록
- `Unique_IP_Port`: IP+Port 기준 최종 중복 제거 후보 목록
- `Run_Summary`: 처리 요약

다른 Application은 폴더명과 `--applications` 값을 바꿔서 같은 방식으로 실행한다.

```bash
python3 tools/extract_ip_port_candidates.py \
  --detail-root ./output/20261007/TCP_OTHER \
  --targets-yaml ./config/detail_targets_filtered.yaml \
  --output ./output/20261007/TCP_OTHER/ip_port_candidates.xlsx \
  --yaml-output ./output/20261007/TCP_OTHER/ip_port_candidates.yaml \
  --applications TCP_OTHER
```

Port 해석:

- Port가 `0`이면 포트를 알 수 없다는 뜻으로 본다.
- Source Port와 Destination Port가 둘 다 `0`이면 확인 필요 row로 분류한다.
- Row Count는 같은 Application / Direction / IP / Port 조합이 Detail Excel에서 몇 줄 나왔는지 나타낸다.
- `Unique_IP_Port`는 Direction이나 중복 row와 관계없이 IP+Port가 같으면 하나로 합친다.

`ip_port_candidates.yaml`에는 `unique_ip_ports`만 저장한다. hostname 조회는 이 목록에서 IP만 중복 제거해서 수행하고, 등록/검토는 IP+Port 기준으로 확인한다.

```yaml
unique_ip_ports:
  - ip: 192.0.2.10
    port: 443
    applications:
      - HTTPS
      - TCP_OTHER
    candidate_count: 3
    total_bitrate: 1000.0
```

## 날짜 폴더 아래 Application별 후보 파일 일괄 생성

Application별 Detail Excel이 이미 생성되어 있다면, 날짜 폴더를 기준으로 각 Application 폴더에 후보 파일을 한 번에 만들 수 있다.
`--targets-yaml`을 지정하면 `detail_targets_filtered.yaml`에서 `enabled: true`인 Application 폴더만 처리한다.

Ubuntu:

```bash
python3 tools/postprocess_by_app.py \
  --date-root ./output/20261007 \
  --targets-yaml ./config/detail_targets_filtered.yaml
```

Windows PowerShell:

```powershell
python .\tools\postprocess_by_app.py `
  --date-root .\output\20261007 `
  --targets-yaml .\config\detail_targets_filtered.yaml
```

먼저 확인만 하려면:

```bash
python3 tools/postprocess_by_app.py \
  --date-root ./output/20261007 \
  --targets-yaml ./config/detail_targets_filtered.yaml \
  --dry-run
```

특정 Application만 처리:

```bash
python3 tools/postprocess_by_app.py \
  --date-root ./output/20261007 \
  --targets-yaml ./config/detail_targets_filtered.yaml \
  --applications HTTPS,TCP_OTHER
```

생성 파일:

```text
output/20261007/HTTPS/ip_port_candidates.xlsx
output/20261007/HTTPS/ip_port_candidates.yaml
output/20261007/TCP_OTHER/ip_port_candidates.xlsx
output/20261007/TCP_OTHER/ip_port_candidates.yaml
```

## IP 후보로 hostname 조회

`extract_ip_port_candidates.py`가 만든 `ip_port_candidates.yaml`의 `unique_ip_ports`를 읽어 hostname 조회 사이트를 호출한다.
설정 파일의 `sites`에는 사이트별 조회/파싱 방식을 지정한다.

지원 방식:

- `type: table`: HTML table의 td 텍스트를 `hostname_index`, `ip_index`로 읽는다.
- `type: input`: 같은 tr 안의 `<input name="IP" value="...">`, `<input name="HOSTNAME" value="...">`를 읽는다.

예:

```yaml
sites:
  - name: Site1
    type: table
    url: "http://192.0.2.10/search"
    search_param_name: "ip"
    hostname_index: 1
    ip_index: 3
  - name: Site2
    type: input
    url: "http://192.0.2.10/search"
    search_param_name: "ip"
```
YAML 입력 시에는 같은 IP가 여러 Port에 있어도 IP 기준으로 한 번만 조회한다.

먼저 설정 파일을 만든다.

Ubuntu:

```bash
cp ./config/hostname_lookup.example.yaml ./config/hostname_lookup.yaml
vim ./config/hostname_lookup.yaml
```

Windows PowerShell:

```powershell
Copy-Item .\config\hostname_lookup.example.yaml .\config\hostname_lookup.yaml
notepad .\config\hostname_lookup.yaml
```

조회 실행:

```bash
python3 tools/lookup_hostnames.py \
  --config ./config/hostname_lookup.yaml \
  --input ./output/20261007/HTTPS/ip_port_candidates.yaml \
  --output ./output/20261007/HTTPS/hostname_lookup.xlsx
```

hostname 조회 결과 Excel의 컬럼 구조:

- `IP`
- `Port`
- `Application`
- `Found`
- `Hostname`
- `Site`
- `Source URL`
- `Error`

Windows PowerShell:

```powershell
python .\tools\lookup_hostnames.py `
  --config .\config\hostname_lookup.yaml `
  --input .\output\20261007\HTTPS\ip_port_candidates.yaml `
  --output .\output\20261007\HTTPS\hostname_lookup.xlsx
```

주의:

- hostname 조회 사이트 URL은 `hostname_lookup.yaml`에만 작성한다.
- 부분 IP 매칭은 하지 않고, table/input 안의 IP가 조회 IP와 정확히 같은 경우만 매칭한다.
- `Site` 컬럼은 hostname이 어느 조회 사이트에서 발견됐는지 표시한다.

## Hostname Docker Lab

고객사 hostname 조회 사이트를 호출하기 전, 로컬 HTTPS lab으로 조회 흐름을 확인할 수 있다.

이 lab은 NETSCOUT API mock이 아니다. NETSCOUT 조회는 실제 테스트 서버를 사용하고, lab은 hostname 조회 사이트만 흉내낸다.

```bash
cd docker/hostname-lab
docker compose up --build
```

조회:

샘플 후보 파일로 먼저 확인:

```bash
python3 tools/lookup_hostnames.py \
  --input ./docker/hostname-lab/sample_ip_port_candidates.yaml \
  --config ./config/hostname_lookup.lab.yaml \
  --output ./output/hostname_lookup_lab_sample.xlsx
```

실제 후보 파일로 확인:

```bash
python3 tools/lookup_hostnames.py \
  --input ./output/20261007/HTTPS/ip_port_candidates.yaml \
  --config ./config/hostname_lookup.lab.yaml \
  --output ./output/20261007/HTTPS/hostname_lookup_lab.xlsx
```

Windows PowerShell:

```powershell
python .\tools\lookup_hostnames.py `
  --input .\output\20261007\HTTPS\ip_port_candidates.yaml `
  --config .\config\hostname_lookup.lab.yaml `
  --output .\output\20261007\HTTPS\hostname_lookup_lab.xlsx
```


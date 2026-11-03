# netscout-auto v0.89 운영 가이드

v0.89의 Inventory, Top N, Detail, 후보 추출 및 Hostname 조회는 `python -m netscout`로 실행합니다. Application 등록 동기화는 `tools/registration`의 독립 도구로 실행합니다.

## 1. 출력 구조

```text
output/YYYYMMDD/
  00_inventory/
  01_top/
  02_targets/
  03_detail/<Application>/
  04_candidates/<Application>/
  05_hostname/<Application>/
  run_summary.txt
```

등록 도구 결과는 별도로 `results/sync_<timestamp>/`에 저장됩니다.

## 2. 기본 운영 순서

1. Inventory 조회
2. Top N 조회
3. Detail 대상 YAML 생성
4. 필요 시 pure/current Inventory 비교로 수동 등록 Application 목록 생성
5. 필요 시 `detail_targets_filtered.yaml` 생성
6. YAML 검토 및 `enabled` 수정
7. Detail Batch 실행
8. IP/Port 후보 생성
9. Hostname 조회
10. `tools/registration`으로 등록 Excel 생성 및 계획 검토
11. 필요한 경우에만 `--apply`와 `SYNC`로 독립 등록 도구 실행

등록 도구의 기본 계획 실행은 읽기 전용이며 삭제는 지원하지 않습니다.

```powershell
python .\tools\registration\make_registration.py --input .\hostname_lookup.xlsx --output .\registration.xlsx
$env:NETSCOUT_PASSWORD = Read-Host "NETSCOUT password"
python .\tools\registration\netscout_sync.py --input .\registration.xlsx --base-url https://NETSCOUT:8443 --username dummy-user --ca-file .\internal-ca.pem
# 실제 반영할 때만 위 명령 끝에 --apply 추가
```

## 3. Inventory

```bash
python3 -m netscout inventory --config ./config/config.test.yaml --env ./.env.local --output-dir ./output
```

## 4. Top N

```bash
python3 -m netscout top --config ./config/config.test.yaml --env ./.env.local --device "nG1vstream635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir ./output
```

## 5. Detail 대상 YAML 생성

```bash
python3 -m netscout detail-targets-yaml --top-xlsx ./output/YYYYMMDD/01_top/TopN파일.xlsx --detail-top-n 50 --limit 1000
```

기본 저장 위치:

```text
output/YYYYMMDD/02_targets/detail_targets.yaml
```

## 6. Detail Batch 실행

```bash
python3 -m netscout run-detail-targets --config ./config/config.test.yaml --env ./.env.local --targets ./output/YYYYMMDD/02_targets/detail_targets.yaml --output-dir ./output --default-duration LAST_7_DAYS
```

## 7. 수동 등록 Application 제외

순정 Inventory와 현재 Inventory를 비교해 수동 등록 Application YAML을 만듭니다.

```bash
python3 -m netscout compare-inventory-apps --pure-inventory ./pure_xlsx/pure_inventory.xlsx --current-inventory ./output/YYYYMMDD/00_inventory/Netscout_Inventory_YYYYMMDD_HHMMSS.xlsx --output ./output/YYYYMMDD/02_targets/manual_registered_apps.csv --yaml-output ./output/YYYYMMDD/02_targets/manual_registered_apps.yaml
```

Detail 대상 YAML에서 수동 등록 Application을 `enabled: false` 처리합니다.

```bash
python3 -m netscout filter-detail-targets --top-yaml ./output/YYYYMMDD/02_targets/detail_targets.yaml --manual-apps-yaml ./output/YYYYMMDD/02_targets/manual_registered_apps.yaml --output ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml
```

기본 보호 대상:

- TCP_OTHER
- UDP_OTHER
- IP_OTHER
- OTHER
- UNKNOWN
- IPV6

## 8. 후보 생성

```bash
python3 -m netscout postprocess --date-root ./output/YYYYMMDD --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml
```

특정 Application만:

```bash
python3 -m netscout postprocess --date-root ./output/YYYYMMDD --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets.yaml --applications HTTPS,TCP_OTHER
```

## 9. 후보 생성과 Hostname 조회를 한 번에 실행

```bash
python3 -m netscout postprocess --date-root ./output/YYYYMMDD --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml --applications HTTPS --hostname-lookup --hostname-config ./config/hostname_lookup.yaml
```

## 10. Hostname 조회만 단독 실행

```bash
python3 -m netscout hostname-lookup --input ./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.yaml --config ./config/hostname_lookup.yaml
```

## 11. 단일 Detail 폴더 후보 추출

```bash
python3 -m netscout extract-candidates --detail-root ./output/YYYYMMDD/03_detail/HTTPS --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets_filtered.yaml --output ./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.xlsx --yaml-output ./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.yaml --applications HTTPS
```

## 12. Windows 예시

```powershell
python -m netscout inventory --config .\config\config.test.yaml --env .\.env.local --output-dir .\output
python -m netscout top --config .\config\config.test.yaml --env .\.env.local --device "nG1vstream635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir .\output
python -m netscout detail-targets-yaml --top-xlsx .\output\YYYYMMDD\01_top\TopN파일.xlsx --detail-top-n 50 --limit 1000
python -m netscout compare-inventory-apps --pure-inventory .\pure_xlsx\pure_inventory.xlsx --current-inventory .\output\YYYYMMDD\00_inventory\Netscout_Inventory_YYYYMMDD_HHMMSS.xlsx --output .\output\YYYYMMDD\02_targets\manual_registered_apps.csv --yaml-output .\output\YYYYMMDD\02_targets\manual_registered_apps.yaml
python -m netscout filter-detail-targets --top-yaml .\output\YYYYMMDD\02_targets\detail_targets.yaml --manual-apps-yaml .\output\YYYYMMDD\02_targets\manual_registered_apps.yaml --output .\output\YYYYMMDD\02_targets\detail_targets_filtered.yaml
python -m netscout run-detail-targets --config .\config\config.test.yaml --env .\.env.local --targets .\output\YYYYMMDD\02_targets\detail_targets_filtered.yaml --output-dir .\output --default-duration LAST_7_DAYS
python -m netscout postprocess --date-root .\output\YYYYMMDD --targets-yaml .\output\YYYYMMDD\02_targets\detail_targets_filtered.yaml
python -m netscout hostname-lookup --input .\output\YYYYMMDD\04_candidates\HTTPS\ip_port_candidates.yaml --config .\config\hostname_lookup.yaml
```

## 13. 보안 체크

- `.env.local`은 공유하지 않습니다.
- 비밀번호, token, cookie는 문서/로그/Excel에 기록하지 않습니다.
- 운영 서버 호출 전에는 반드시 승인받습니다.

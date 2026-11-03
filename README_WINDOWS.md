# netscout-auto v0.89 Windows 사용 가이드

Windows PowerShell에서는 `netscout.py`가 있는 `netscout-auto` 폴더에서 실행합니다.

## 준비

```powershell
cd D:\path\to\netscout-auto
Copy-Item .\config\config.example.yaml .\config\config.test.yaml
Copy-Item .\.env.local.example .\.env.local
notepad .\config\config.test.yaml
notepad .\.env.local
python -m netscout --help
```

## Inventory

```powershell
python -m netscout inventory --config .\config\config.test.yaml --env .\.env.local --output-dir .\output
```

## Application 등록 동기화

```powershell
python .\tools\registration\make_registration.py --input .\hostname_lookup.xlsx --output .\registration.xlsx
$env:NETSCOUT_PASSWORD = Read-Host "NETSCOUT password"
python .\tools\registration\netscout_sync.py --input .\registration.xlsx --base-url https://NETSCOUT:8443 --username dummy-user --ca-file .\internal-ca.pem
# 계획을 검토한 뒤 실제 반영이 필요할 때만 마지막에 --apply 추가
```

기본 실행은 조회와 계획 생성만 수행합니다. `--apply` 실행은 계획 확인 후 `SYNC`를 정확히 입력해야 POST/PUT을 수행합니다. 테스트용 자체 서명 인증서라면 `--ca-file` 대신 `--insecure`를 사용할 수 있습니다.

## Top N

```powershell
python -m netscout top --config .\config\config.test.yaml --env .\.env.local --device "nG1vstream635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir .\output
```

## Detail 대상 YAML

```powershell
python -m netscout detail-targets-yaml --top-xlsx .\output\YYYYMMDD\01_top\TopN파일.xlsx --detail-top-n 50 --limit 1000
```

## Detail Batch

```powershell
python -m netscout run-detail-targets --config .\config\config.test.yaml --env .\.env.local --targets .\output\YYYYMMDD\02_targets\detail_targets.yaml --output-dir .\output --default-duration LAST_7_DAYS
```

## IP/Port 후보 생성

```powershell
python -m netscout postprocess --date-root .\output\YYYYMMDD --targets-yaml .\output\YYYYMMDD\02_targets\detail_targets.yaml
```

## Hostname 조회 포함 실행

```powershell
python -m netscout postprocess --date-root .\output\YYYYMMDD --targets-yaml .\output\YYYYMMDD\02_targets\detail_targets.yaml --applications HTTPS --hostname-lookup --hostname-config .\config\hostname_lookup.yaml
```

## Hostname 조회만 실행

```powershell
python -m netscout hostname-lookup --input .\output\YYYYMMDD\04_candidates\HTTPS\ip_port_candidates.yaml --config .\config\hostname_lookup.yaml
```

## 단일 Detail 폴더 후보 추출

```powershell
python -m netscout extract-candidates --detail-root .\output\YYYYMMDD\03_detail\HTTPS --targets-yaml .\output\YYYYMMDD\02_targets\detail_targets.yaml --output .\output\YYYYMMDD\04_candidates\HTTPS\ip_port_candidates.xlsx --yaml-output .\output\YYYYMMDD\04_candidates\HTTPS\ip_port_candidates.yaml --applications HTTPS
```

## 결과 구조

```text
output\YYYYMMDD\
  00_inventory\
  01_top\
  02_targets\
  03_detail\
  04_candidates\
  05_hostname\
  run_summary.txt
```

## 보안

- `.env.local`은 공유하지 않습니다.
- 비밀번호, token, cookie는 문서/로그/Excel에 기록하지 않습니다.

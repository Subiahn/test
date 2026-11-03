# netscout-auto v0.89 Ubuntu 사용 가이드

Ubuntu에서는 `netscout.py`가 있는 `netscout-auto` 폴더에서 실행합니다.

## 준비

```bash
cd ~/netscout-auto
cp ./config/config.example.yaml ./config/config.test.yaml
cp ./.env.local.example ./.env.local
vim ./config/config.test.yaml
vim ./.env.local
python3 -m netscout --help
```

경로는 Linux 형식인 `./config/config.test.yaml`을 사용합니다.

## Inventory

```bash
python3 -m netscout inventory --config ./config/config.test.yaml --env ./.env.local --output-dir ./output
```

## Application 등록 동기화

```bash
python3 ./tools/registration/make_registration.py --input ./hostname_lookup.xlsx --output ./registration.xlsx
read -s NETSCOUT_PASSWORD && export NETSCOUT_PASSWORD
python3 ./tools/registration/netscout_sync.py --input ./registration.xlsx --base-url https://NETSCOUT:8443 --username dummy-user --ca-file ./internal-ca.pem
# 계획을 검토한 뒤 실제 반영이 필요할 때만 마지막에 --apply 추가
```

기본 실행은 읽기 전용이며 실제 적용에는 `--apply`와 `SYNC` 확인이 필요합니다.

## Top N

```bash
python3 -m netscout top --config ./config/config.test.yaml --env ./.env.local --device "nG1vstream635" --interface "if3" --duration LAST_7_DAYS --top-n 50 --output-dir ./output
```

고정 시간 범위:

```bash
python3 -m netscout top --config ./config/config.test.yaml --env ./.env.local --device "nG1vstream635" --interface "if3" --start-time 2026-10-07_00:00:00 --end-time 2026-10-14_00:00:00 --resolution NO_RESOLUTION --top-n 50 --output-dir ./output
```

## Detail 대상 YAML

```bash
python3 -m netscout detail-targets-yaml --top-xlsx ./output/YYYYMMDD/01_top/TopN파일.xlsx --detail-top-n 50 --limit 1000
```

## Detail Batch

Dry-run:

```bash
python3 -m netscout run-detail-targets --config ./config/config.test.yaml --env ./.env.local --targets ./output/YYYYMMDD/02_targets/detail_targets.yaml --output-dir ./output --default-duration LAST_7_DAYS --dry-run
```

실행:

```bash
python3 -m netscout run-detail-targets --config ./config/config.test.yaml --env ./.env.local --targets ./output/YYYYMMDD/02_targets/detail_targets.yaml --output-dir ./output --default-duration LAST_7_DAYS
```

## IP/Port 후보 생성

```bash
python3 -m netscout postprocess --date-root ./output/YYYYMMDD --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets.yaml
```

## Hostname 조회 포함 실행

```bash
python3 -m netscout postprocess --date-root ./output/YYYYMMDD --targets-yaml ./output/YYYYMMDD/02_targets/detail_targets.yaml --applications HTTPS --hostname-lookup --hostname-config ./config/hostname_lookup.yaml
```

## Hostname 조회만 실행

```bash
python3 -m netscout hostname-lookup --input ./output/YYYYMMDD/04_candidates/HTTPS/ip_port_candidates.yaml --config ./config/hostname_lookup.yaml
```

## 결과 구조

```text
output/YYYYMMDD/
  00_inventory/
  01_top/
  02_targets/
  03_detail/
  04_candidates/
  05_hostname/
  run_summary.txt
```

## 자주 나는 오류

- `No module named netscout`: `netscout.py`가 있는 폴더에서 실행합니다.
- `Config file not found`: Ubuntu에서는 `.\config\...`가 아니라 `./config/...`를 씁니다.
- `Missing NETSCOUT_USERNAME or NETSCOUT_PASSWORD`: `.env.local`을 확인합니다.
- `Permission denied`: Excel 파일이 열려 있거나 output 권한이 없는지 확인합니다.

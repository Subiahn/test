# netscout-auto v0.87 설계 메모

v0.87은 v0.86의 날짜별 output 구조를 유지하면서, 후처리 tools를 `python -m netscout` CLI로 통합한 버전이다.

## v0.86 대비 변경점

- `postprocess` 명령 추가
- `extract-candidates` 명령 추가
- `hostname-lookup` 명령 추가
- `postprocess --hostname-lookup`으로 후보 생성과 Hostname 조회를 연속 실행
- 기존 tools 스크립트는 내부 모듈로 재사용

## 출력 구조

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

## 통합 CLI 매핑

```text
tools/postprocess_by_app.py
  -> python -m netscout postprocess

tools/extract_ip_port_candidates.py
  -> python -m netscout extract-candidates

tools/lookup_hostnames.py
  -> python -m netscout hostname-lookup
```

## 권장 흐름

```bash
python3 -m netscout inventory ...
python3 -m netscout top ...
python3 -m netscout detail-targets-yaml ...
python3 -m netscout run-detail-targets ...
python3 -m netscout postprocess ...
python3 -m netscout hostname-lookup ...
```


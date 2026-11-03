# NETSCOUT Application 등록 도구

이 디렉터리는 v0.89 기반의 `python -m netscout` CLI와 분리된 독립 도구입니다. 기존 CLI와 `src/netscout_automation/client.py`를 변경하지 않습니다.

## 구성

- `make_registration.py`: Hostname 조회 Excel을 검토용 Registration Excel로 변환합니다. API를 호출하지 않습니다.
- `netscout_sync.py`: 현재 Application을 GET으로 조회해 계획을 만들고, 선택적으로 적용합니다.
- `registration_rules.py`: 입력값 검증과 XML 생성 규칙입니다.
- `sync_planner.py`: `CREATE`, `MATCHED`, `ADD_IP`, `ADD_PORT`, `CONFLICT`를 판정합니다.
- `netscout_api.py`: 등록 도구 전용 Config API 클라이언트입니다.
- `tests/`: 등록 도구 전용 자동 테스트입니다.

## 1. 등록 Excel 만들기

프로젝트 루트에서 실행합니다.

```powershell
python .\tools\registration\make_registration.py `
  --input .\hostname_lookup.xlsx `
  --output .\registration.xlsx
```

생성된 Excel에서 Application 이름, 유형, 부모, IP, 포트, 프로토콜을 검토하십시오. 이름은 자동 생성하지 않습니다.

## 2. 계획만 확인하기

비밀번호는 명령행 인수나 파일에 저장하지 않고 환경변수로 전달합니다.

```powershell
$env:NETSCOUT_PASSWORD = Read-Host "NETSCOUT password"

python .\tools\registration\netscout_sync.py `
  --input .\registration.xlsx `
  --base-url https://test-server.example.com:8443 `
  --username dummy-user `
  --insecure
```

`--apply`가 없으면 GET 조회와 계획 파일 생성만 수행합니다. 사내 CA 파일이 있으면 `--insecure` 대신 `--ca-file .\internal-ca.pem`을 권장합니다. 사용자 이름/비밀번호 인증과 서버 인증서 검증은 서로 다른 기능이므로, 인증서 검증 설정은 별도로 선택해야 합니다.

판정 의미:

- `CREATE`: 같은 이름의 Application이 없어 신규 생성 대상
- `MATCHED`: 이름, IP, 포트 등 관리 대상 값이 이미 일치
- `ADD_IP`: 기존 Application에 새 IP만 합집합으로 추가
- `ADD_PORT`: 기존 Application에 새 포트만 합집합으로 추가
- `CONFLICT`: 다른 Application과 주소가 충돌하거나 유형/부모 등 변경 불가 조건이 불일치

동일한 파일을 다시 계획하면 서버 상태가 이미 같을 때 `MATCHED`가 됩니다. 같은 이름에 IP만 추가된 경우에는 기존 IP를 유지한 채 `ADD_IP`가 계획됩니다.

## 3. 실제 적용하기

계획 결과와 Excel을 검토한 뒤에만 `--apply`를 추가합니다.

```powershell
python .\tools\registration\netscout_sync.py `
  --input .\registration.xlsx `
  --base-url https://test-server.example.com:8443 `
  --username dummy-user `
  --insecure `
  --apply
```

도구가 보여 주는 계획을 확인한 후 `SYNC`를 정확히 입력해야 적용됩니다. 이 도구는 기존 IP/포트 제거와 Application 삭제를 하지 않습니다. `CONFLICT` 항목도 자동 적용하지 않습니다.

ADD_IP/ADD_PORT의 PUT 요청은 변경 필드만 보내지 않고 검토·승인된 desired 필드 전체에 기존 값과 신규 값의 합집합을 반영합니다. 이는 nGeniusONE v6.3.4 장비가 PUT에서 `ApplicationType` 등 식별 필드를 요구하는 동작을 반영한 것입니다.

결과는 `results/sync_날짜_시간/` 아래의 `inventory_before.json`, `plan.json` 등에 저장됩니다.

## 테스트

```powershell
python -m pytest -q
python -m compileall -q .\tools\registration
```

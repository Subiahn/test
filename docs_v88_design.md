# netscout-auto v8.8 설계 메모

## 경계

v8.8은 v8.7의 `python -m netscout` CLI와 공통 API 클라이언트를 유지합니다. Config API 기반 Application 등록 기능은 결합하지 않고 독립 실행 도구로 격리합니다.

```text
기존 V8 CLI (조회·분석·후처리)
    └─ 변경 없이 유지

tools/registration (Application 등록)
    ├─ Excel 변환
    ├─ 현재 Application GET
    ├─ 변경 계획 생성
    └─ 명시적 --apply + SYNC 확인 시 POST/PUT

tools/nca
    └─ 향후 두 번째 Hostname 호출 모듈을 위한 예약 디렉터리
```

## 등록 도구 모듈

- `make_registration.py`: Excel 변환 전용, API 호출 없음
- `registration_rules.py`: 입력 검증과 XML 계약
- `sync_planner.py`: 현재 상태와 요청 상태 비교
- `netscout_api.py`: 등록 도구 전용 HTTP 처리
- `netscout_sync.py`: 계획/적용 실행 진입점

## 안전 계약

- 기본 동작은 GET 기반 계획 생성입니다.
- 쓰기는 `--apply`와 `SYNC` 확인을 모두 요구합니다.
- 업데이트는 IP/포트 합집합만 허용합니다.
- 삭제와 기존 값 제거는 수행하지 않습니다.
- 충돌은 `CONFLICT`로 표시하고 적용하지 않습니다.
- 인증정보는 환경변수로만 전달합니다.

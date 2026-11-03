# Hostname Lookup Lab

이 Docker lab은 고객사 hostname 조회 사이트를 로컬에서 흉내내기 위한 테스트 서버다.
NETSCOUT API mock 서버가 아니다.

NETSCOUT Application / Top / Detail 조회는 실제 테스트 서버를 사용하고,
이 lab은 `ip_port_candidates.yaml`을 만든 뒤 hostname 조회 단계만 검증한다.

## 구성

- `https://127.0.0.1:5000/search`: HTML table 방식
- `https://127.0.0.1:5002/search`: input value 방식
- 두 포트는 같은 컨테이너에서 실행된다.
- 자체 서명 인증서를 사용하므로 `config/hostname_lookup.lab.yaml`에서는 `verify_ssl: false`를 사용한다.

## 실행

```bash
cd docker/hostname-lab
docker compose up --build
```

Windows PowerShell:

```powershell
cd .\docker\hostname-lab
docker compose up --build
```

## 조회 테스트

실제 `ip_port_candidates.yaml`이 아직 없으면 lab에 포함된 샘플 파일로 먼저 테스트할 수 있다.

```bash
python3 tools/lookup_hostnames.py \
  --input ./docker/hostname-lab/sample_ip_port_candidates.yaml \
  --config ./config/hostname_lookup.lab.yaml \
  --output ./output/hostname_lookup_lab_sample.xlsx
```

실제 Detail 후보 파일을 사용할 때:

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

## 더미 데이터 예

- `192.0.2.10` -> `md-feed-gw-01`
- `192.0.2.10` -> `md-feed-gw-02`
- `192.0.2.10` -> `market-data-router-01`
- `192.0.2.10` -> `order-gw-01`

검색 결과에는 유사 IP도 같이 반환되지만, lookup 도구는 요청 IP와 완전히 같은 IP만 채택한다.

"""`lookup_hostnames.py` 실행 호환용 wrapper.

운영 문서의 정식 파일명은 `lookup_hostnames.py`지만, 사람이 singular 형태인
`lookup_hostname.py`로 실행해도 같은 동작을 하도록 얇게 연결한다.
"""

from __future__ import annotations

from lookup_hostnames import main


if __name__ == "__main__":
    raise SystemExit(main())

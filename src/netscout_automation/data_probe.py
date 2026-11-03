"""Raw Data API probe 전송기.

parser 가정을 확정하기 전에 raw response를 먼저 저장하기 위한 모듈이다.
placeholder validation으로 템플릿 XML이 실수로 실제 NETSCOUT 서버에 전송되는 것을 막는다.
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from time import perf_counter
from urllib.parse import urljoin, urlparse

import requests
import urllib3
from requests.auth import HTTPBasicAuth

from .config import NetscoutSettings

logger = logging.getLogger(__name__)

ACCEPT_HEADERS = {
    "json": "application/json",
    "xml": "application/xml",
    "csv": "text/csv",
}

RAW_EXTENSIONS = {
    "json": ".json",
    "xml": ".xml",
    "csv": ".csv",
}

BLOCKED_PLACEHOLDER_MARKERS = ("TODO_", "PLACEHOLDER", "CHANGE_ME")


class DataProbeError(RuntimeError):
    """Data API probe를 완료할 수 없을 때 발생한다."""


@dataclass(frozen=True)
class ProbeRequestPreview:
    url: str
    host: str
    endpoint: str
    content_type: str
    accept_type: str
    xml_file: Path
    verify_ssl: bool
    timeout: int


@dataclass(frozen=True)
class ProbeResult:
    status_code: int
    elapsed_seconds: float
    content_type: str
    output_path: Path


def load_probe_xml(xml_file: str | Path) -> str:
    path = Path(xml_file)
    if not path.exists():
        raise DataProbeError(f"Probe XML file not found: {path}")
    text = path.read_text(encoding="utf-8")
    if not text.strip():
        raise DataProbeError(f"Probe XML file is empty: {path}")
    return text


def validate_probe_xml_ready_for_request(xml_text: str, xml_file: str | Path) -> None:
    """XML에 템플릿 placeholder가 남아 있으면 실제 API 호출을 차단한다."""
    found = [marker for marker in BLOCKED_PLACEHOLDER_MARKERS if marker in xml_text]
    if found:
        markers = ", ".join(found)
        raise DataProbeError(
            f"Probe XML contains unresolved placeholder marker(s): {markers}. "
            f"Create an actual request XML before calling Data API: {Path(xml_file)}"
        )


def build_probe_preview(settings: NetscoutSettings, xml_file: str | Path, accept: str) -> ProbeRequestPreview:
    accept = _validate_accept(accept)
    url = _url(settings.base_url, settings.data_api_endpoint)
    return ProbeRequestPreview(
        url=url,
        host=urlparse(url).hostname or "unknown-host",
        endpoint=settings.data_api_endpoint,
        content_type="application/xml",
        accept_type=ACCEPT_HEADERS[accept],
        xml_file=Path(xml_file),
        verify_ssl=settings.verify_ssl,
        timeout=settings.timeout,
    )


def timestamped_raw_response_path(output_dir: str | Path, accept: str, now: datetime | None = None) -> Path:
    accept = _validate_accept(accept)
    timestamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    return Path(output_dir) / "raw" / f"data_probe_{timestamp}{RAW_EXTENSIONS[accept]}"


def run_data_probe(
    settings: NetscoutSettings,
    xml_file: str | Path,
    accept: str,
    output_dir: str | Path | None = None,
) -> ProbeResult:
    """사용자가 지정한 XML probe를 전송하고 파싱 없이 raw response를 저장한다."""
    accept = _validate_accept(accept)
    xml_text = load_probe_xml(xml_file)
    validate_probe_xml_ready_for_request(xml_text, xml_file)
    preview = build_probe_preview(settings, xml_file, accept)

    logger.info(
        "Calling Data API probe host=%s endpoint=%s verify_ssl=%s accept=%s",
        preview.host,
        preview.endpoint,
        preview.verify_ssl,
        accept,
    )

    try:
        if not settings.verify_ssl:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        start = perf_counter()
        response = requests.post(
            preview.url,
            data=xml_text.encode("utf-8"),
            auth=HTTPBasicAuth(settings.username, settings.password),
            headers={"Content-Type": preview.content_type, "Accept": preview.accept_type},
            verify=settings.verify_ssl,
            timeout=settings.timeout,
        )
        elapsed = perf_counter() - start
    except requests.exceptions.Timeout as exc:
        raise DataProbeError("Data API probe timed out. Check network and timeout settings.") from exc
    except requests.exceptions.SSLError as exc:
        raise DataProbeError("Data API probe failed with an SSL error. Check verify_ssl and server certificate.") from exc
    except requests.exceptions.ConnectionError as exc:
        raise DataProbeError("Data API probe connection failed. Check base_url and network access.") from exc
    except requests.exceptions.RequestException as exc:
        raise DataProbeError("Data API probe request failed.") from exc

    content_type = response.headers.get("Content-Type", "unknown")
    logger.info(
        "Data API probe response status=%s elapsed=%.3fs content_type=%s",
        response.status_code,
        elapsed,
        content_type,
    )

    if response.status_code != 200:
        # raw error body에는 민감한 서버 문구가 포함될 수 있으므로 짧고 마스킹된 preview만 남긴다.
        preview_text = _sanitized_preview(response.text)
        raise DataProbeError(
            f"Data API probe failed with HTTP {response.status_code}. "
            f"Content-Type={content_type}. Response preview={preview_text!r}"
        )

    path = timestamped_raw_response_path(output_dir or settings.output_directory, accept)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(response.content)
    logger.info("Saved Data API raw response to %s", path)

    return ProbeResult(
        status_code=response.status_code,
        elapsed_seconds=elapsed,
        content_type=content_type,
        output_path=path,
    )


def _url(base_url: str, endpoint: str) -> str:
    return urljoin(base_url.rstrip("/") + "/", endpoint.lstrip("/"))


def _validate_accept(accept: str) -> str:
    normalized = accept.lower()
    if normalized not in ACCEPT_HEADERS:
        raise DataProbeError(f"Unsupported accept type: {accept}")
    return normalized


def _sanitized_preview(text: str, limit: int = 300) -> str:
    compact = " ".join((text or "").split())
    try:
        parsed = json.loads(compact)
    except json.JSONDecodeError:
        return compact[:limit]
    if isinstance(parsed, dict):
        parsed = _mask_sensitive_dict(parsed)
    return json.dumps(parsed, ensure_ascii=True)[:limit]


def _mask_sensitive_dict(value: dict) -> dict:
    masked = {}
    for key, item in value.items():
        if key.lower() in {"password", "token", "cookie", "authorization", "apikey", "api_key"}:
            masked[key] = "***"
        elif isinstance(item, dict):
            masked[key] = _mask_sensitive_dict(item)
        else:
            masked[key] = item
    return masked


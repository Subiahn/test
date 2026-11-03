"""NETSCOUT API CLI 명령에서 사용하는 설정과 인증정보를 로드한다.

환경별 값은 YAML에서 읽고, username/password 또는 NSSESSIONID token은
`.env.local` 또는 Windows 환경변수에서 읽는다. 비밀번호/token이 코드,
문서, 설정 템플릿, 로그에 남지 않도록 분리한다.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from .auth import is_token_auth, normalized_auth_type


@dataclass(frozen=True)
class NetscoutSettings:
    """Config API와 Data API client가 함께 사용하는 실행 시점 설정값."""

    base_url: str
    auth_type: str
    verify_ssl: bool
    timeout: int
    config_applications_endpoint: str
    data_api_endpoint: str
    output_directory: str
    username: str
    password: str
    data_auth_type: str = "basic"
    session_token: str = ""
    session_token_endpoint: str = ""
    session_token_file: str = ""
    session_token_method: str = "GET"
    session_token_content_type: str = ""
    session_token_body: str = ""
    session_token_cookie_name: str = "NSSESSIONID"
    devices_endpoint: str = "/ng1api/ncm/devices"
    device_interfaces_endpoint_template: str = "/ng1api/ncm/devices/{device_name}/interfaces"


class ConfigError(ValueError):
    """설정값 또는 필수 인증정보가 없을 때 발생한다."""


def load_dotenv_file(path: str | Path = ".env.local") -> None:
    """기존 환경변수를 덮어쓰지 않고 단순 KEY=VALUE 형식만 로드한다."""
    env_path = Path(path)
    if not env_path.exists():
        return

    for line in env_path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def load_settings(config_path: str | Path, env_path: str | Path = ".env.local") -> NetscoutSettings:
    """CLI 1회 실행에 필요한 YAML 설정과 NETSCOUT 인증정보를 로드한다."""
    load_dotenv_file(env_path)

    path = Path(config_path)
    if not path.exists():
        raise ConfigError(f"Config file not found: {path}")

    try:
        data: dict[str, Any] = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as exc:
        raise ConfigError(f"Invalid YAML config: {path}") from exc

    netscout = data.get("netscout") or {}
    api = data.get("api") or {}
    output = data.get("output") or {}

    auth_type = normalized_auth_type(netscout.get("auth_type", "basic"))
    data_auth_type = normalized_auth_type(netscout.get("data_auth_type", auth_type))
    username = os.getenv("NETSCOUT_USERNAME", "")
    password = os.getenv("NETSCOUT_PASSWORD", "")
    session_token = first_env_value(
        "NETSCOUT_NSSESSIONID",
        "NETSCOUT_SESSION_TOKEN",
        "NETSCOUT_API_TOKEN",
        "NETSCOUT_TOKEN",
    )
    # 인증정보는 메모리로만 읽는다. password/token은 YAML 파일이나 로그에 남기지 않는다.
    if auth_type == "basic" and (not username or not password):
        raise ConfigError(
            "Missing NETSCOUT_USERNAME or NETSCOUT_PASSWORD. Set them in .env.local or Windows environment variables."
        )
    if is_token_auth(auth_type) or is_token_auth(data_auth_type):
        if not session_token:
            raise ConfigError(
                "Missing NETSCOUT_NSSESSIONID. Set it in .env.local or environment variables for token auth."
            )
    elif not username or not password:
        raise ConfigError(
            "Missing NETSCOUT_USERNAME or NETSCOUT_PASSWORD. Set them in .env.local or Windows environment variables."
        )

    return NetscoutSettings(
        base_url=_required(netscout, "base_url"),
        auth_type=auth_type,
        data_auth_type=data_auth_type,
        verify_ssl=bool(netscout.get("verify_ssl", True)),
        timeout=int(netscout.get("timeout", 30)),
        config_applications_endpoint=_required(api, "config_applications_endpoint"),
        data_api_endpoint=api.get("data_api_endpoint") or api.get("data_tcp_other_endpoint") or "/dbonequerydata",
        output_directory=output.get("directory", "output"),
        username=username,
        password=password,
        session_token=session_token,
        session_token_endpoint=api.get("session_token_endpoint", ""),
        session_token_file=netscout.get("session_token_file", ""),
        session_token_method=str(api.get("session_token_method", "GET")).upper(),
        session_token_content_type=api.get("session_token_content_type", ""),
        session_token_body=api.get("session_token_body", ""),
        session_token_cookie_name=api.get("session_token_cookie_name", "NSSESSIONID"),
        devices_endpoint=api.get("devices_endpoint", "/ng1api/ncm/devices"),
        device_interfaces_endpoint_template=api.get(
            "device_interfaces_endpoint_template",
            "/ng1api/ncm/devices/{device_name}/interfaces",
        ),
    )


def _required(section: dict[str, Any], key: str) -> str:
    value = section.get(key)
    if value in (None, ""):
        raise ConfigError(f"Missing required config value: {key}")
    return str(value)


def first_env_value(*keys: str) -> str:
    """여러 환경변수 후보 중 첫 번째 값을 반환한다."""
    for key in keys:
        value = os.getenv(key, "")
        if value:
            return value
    return ""

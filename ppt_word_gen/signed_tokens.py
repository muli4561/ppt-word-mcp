# coding=utf-8
"""短效 HMAC 令牌：用于产物下载和一次性上传票据。"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import threading
import time
from pathlib import Path
from typing import Any, Dict

from .config import (
    DOWNLOAD_SIGNING_SECRET,
    DOWNLOAD_SIGNING_SECRET_FILE,
    SIGNED_URL_EXPIRE_SECONDS,
)


class InvalidSignedToken(ValueError):
    """令牌格式、签名、类型或有效期不正确。"""


_secret_lock = threading.Lock()
_cached_secret: bytes | None = None


def _b64encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _b64decode(value: str) -> bytes:
    try:
        return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))
    except Exception as exc:  # noqa: BLE001
        raise InvalidSignedToken("令牌 Base64 无效") from exc


def _secret() -> bytes:
    """优先使用环境密钥；否则创建可跨容器重启复用的本地密钥文件。"""
    global _cached_secret
    if _cached_secret is not None:
        return _cached_secret
    with _secret_lock:
        if _cached_secret is not None:
            return _cached_secret
        if DOWNLOAD_SIGNING_SECRET:
            value = DOWNLOAD_SIGNING_SECRET.encode("utf-8")
        else:
            path = Path(DOWNLOAD_SIGNING_SECRET_FILE)
            path.parent.mkdir(parents=True, exist_ok=True)
            try:
                value = path.read_bytes().strip()
            except FileNotFoundError:
                value = secrets.token_hex(32).encode("ascii")
                temporary = path.with_suffix(path.suffix + ".part")
                temporary.write_bytes(value)
                try:
                    os.chmod(temporary, 0o600)
                except OSError:
                    pass
                temporary.replace(path)
        if len(value) < 32:
            raise RuntimeError("下载签名密钥至少需要 32 字节")
        _cached_secret = value
        return value


def sign_claims(kind: str, claims: Dict[str, Any], ttl_seconds: int | None = None) -> str:
    now = int(time.time())
    payload = {
        "v": 1,
        "kind": kind,
        "iat": now,
        "exp": now + int(ttl_seconds or SIGNED_URL_EXPIRE_SECONDS),
        **claims,
    }
    encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode(
        "utf-8"
    )
    signature = hmac.new(_secret(), encoded, hashlib.sha256).digest()
    return f"{_b64encode(encoded)}.{_b64encode(signature)}"


def verify_claims(token: str, expected_kind: str) -> Dict[str, Any]:
    try:
        payload_part, signature_part = token.split(".", 1)
    except ValueError as exc:
        raise InvalidSignedToken("令牌格式无效") from exc
    payload_bytes = _b64decode(payload_part)
    supplied = _b64decode(signature_part)
    expected = hmac.new(_secret(), payload_bytes, hashlib.sha256).digest()
    if not hmac.compare_digest(supplied, expected):
        raise InvalidSignedToken("令牌签名无效")
    try:
        payload = json.loads(payload_bytes)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise InvalidSignedToken("令牌内容无效") from exc
    if not isinstance(payload, dict) or payload.get("kind") != expected_kind:
        raise InvalidSignedToken("令牌用途不匹配")
    if int(payload.get("exp", 0)) < int(time.time()):
        raise InvalidSignedToken("令牌已过期")
    return payload


def expires_at(token: str, expected_kind: str) -> int:
    return int(verify_claims(token, expected_kind)["exp"])

"""Shared inline artifact response packaging for Office MCPs.

This module does not render artifacts. It packages already-rendered bytes or
files into a consistent response shape for consolidated export dispatchers.
"""
from __future__ import annotations

import base64
from pathlib import Path
from typing import Any

DEFAULT_MAX_INLINE_BYTES = 3_000_000
SUPPORTED_RESPONSE_MODES: frozenset[str] = frozenset(
    {"auto", "content_block", "base64_json"}
)
DEFAULT_ALLOWED_MIME_TYPES: frozenset[str] = frozenset({"image/png"})


def validate_inline_artifact_options(
    *,
    mime_type: str,
    response_mode: str,
    encoded_or_raw_size: int | None,
    max_inline_bytes: int,
    allowed_mime_types: set[str] | frozenset[str],
) -> None:
    """Validate common inline artifact options.

    Size overages are not validation failures because callers should still be
    able to return a successful file export with explicit omission metadata.
    """
    if response_mode not in SUPPORTED_RESPONSE_MODES:
        raise ValueError(
            "response_mode must be one of: auto, content_block, base64_json"
        )
    if max_inline_bytes <= 0:
        raise ValueError("max_inline_bytes must be greater than 0")
    if not mime_type:
        raise ValueError("mime_type is required")
    if not allowed_mime_types:
        raise ValueError("allowed_mime_types must not be empty")
    if encoded_or_raw_size is not None and encoded_or_raw_size < 0:
        raise ValueError("encoded_or_raw_size must be non-negative")


def _omitted_artifact(
    *,
    artifact_type: str,
    mime_type: str,
    omitted_reason: str,
    encoded_bytes: int | None,
    max_inline_bytes: int,
    metadata: dict[str, Any],
) -> dict[str, Any]:
    artifact: dict[str, Any] = {
        "artifact_type": artifact_type,
        "mime_type": mime_type,
        "inline": False,
        "omitted_reason": omitted_reason,
        "max_inline_bytes": max_inline_bytes,
        "metadata": metadata,
    }
    if encoded_bytes is not None:
        artifact["encoded_bytes"] = encoded_bytes
    return artifact


def build_inline_artifact_response(
    *,
    artifact_type: str,
    mime_type: str,
    bytes_data: bytes | None = None,
    output_path: str | None = None,
    return_inline: bool = False,
    response_mode: str = "auto",
    max_inline_bytes: int = DEFAULT_MAX_INLINE_BYTES,
    metadata: dict | None = None,
    allowed_mime_types: set[str] | frozenset[str] = DEFAULT_ALLOWED_MIME_TYPES,
) -> dict[str, Any]:
    """Build a standard inline artifact response.

    If ``return_inline`` is false, no payload is included. If payload bytes are
    too large or unsupported, the underlying export can still be reported as
    successful through ``output_path`` with omission metadata.
    """
    validate_inline_artifact_options(
        mime_type=mime_type,
        response_mode=response_mode,
        encoded_or_raw_size=len(bytes_data) if bytes_data is not None else None,
        max_inline_bytes=max_inline_bytes,
        allowed_mime_types=allowed_mime_types,
    )
    artifact_metadata: dict[str, Any] = dict(metadata or {})
    response: dict[str, Any] = {"output_path": output_path}

    if not return_inline:
        response["artifact"] = _omitted_artifact(
            artifact_type=artifact_type,
            mime_type=mime_type,
            omitted_reason="inline disabled",
            encoded_bytes=None,
            max_inline_bytes=max_inline_bytes,
            metadata=artifact_metadata,
        )
        return response

    if mime_type not in allowed_mime_types:
        response["artifact"] = _omitted_artifact(
            artifact_type=artifact_type,
            mime_type=mime_type,
            omitted_reason="unsupported mime type",
            encoded_bytes=None,
            max_inline_bytes=max_inline_bytes,
            metadata=artifact_metadata,
        )
        return response

    payload = bytes_data
    if payload is None and output_path is not None:
        payload = Path(output_path).read_bytes()

    if not payload:
        response["artifact"] = _omitted_artifact(
            artifact_type=artifact_type,
            mime_type=mime_type,
            omitted_reason="missing artifact bytes",
            encoded_bytes=0 if payload == b"" else None,
            max_inline_bytes=max_inline_bytes,
            metadata=artifact_metadata,
        )
        return response

    encoded = base64.b64encode(payload).decode("ascii")
    encoded_bytes = len(encoded.encode("ascii"))
    if encoded_bytes > max_inline_bytes:
        response["artifact"] = _omitted_artifact(
            artifact_type=artifact_type,
            mime_type=mime_type,
            omitted_reason="encoded image exceeds max_inline_bytes",
            encoded_bytes=encoded_bytes,
            max_inline_bytes=max_inline_bytes,
            metadata=artifact_metadata,
        )
        return response

    if response_mode == "content_block":
        response["artifact"] = _omitted_artifact(
            artifact_type=artifact_type,
            mime_type=mime_type,
            omitted_reason="unsupported response mode",
            encoded_bytes=encoded_bytes,
            max_inline_bytes=max_inline_bytes,
            metadata=artifact_metadata,
        )
        return response

    response["artifact"] = {
        "artifact_type": artifact_type,
        "mime_type": mime_type,
        "inline": True,
        "delivery": "base64_json",
        "encoded_bytes": encoded_bytes,
        "image_base64": encoded,
        "metadata": artifact_metadata,
    }
    return response

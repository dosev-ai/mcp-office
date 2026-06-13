from __future__ import annotations

import base64
import logging

import pytest

from mcpshared.inline_artifact_response import (
    build_inline_artifact_response,
    validate_inline_artifact_options,
)


PNG_BYTES = b"\x89PNG\r\n\x1a\nfake-png"


def test_return_inline_false_omits_payload(tmp_path):
    out = tmp_path / "slide.png"
    out.write_bytes(PNG_BYTES)

    result = build_inline_artifact_response(
        artifact_type="slide_render",
        mime_type="image/png",
        output_path=str(out),
        return_inline=False,
    )

    assert result["output_path"] == str(out)
    assert result["artifact"]["inline"] is False
    assert result["artifact"]["omitted_reason"] == "inline disabled"
    assert "image_base64" not in result["artifact"]


def test_valid_png_payload_base64_json_decodes():
    result = build_inline_artifact_response(
        artifact_type="slide_render",
        mime_type="image/png",
        bytes_data=PNG_BYTES,
        return_inline=True,
        response_mode="base64_json",
    )

    artifact = result["artifact"]
    assert artifact["inline"] is True
    assert artifact["delivery"] == "base64_json"
    assert base64.b64decode(artifact["image_base64"]) == PNG_BYTES
    assert artifact["encoded_bytes"] == len(artifact["image_base64"].encode("ascii"))


def test_auto_mode_uses_base64_json_initially():
    result = build_inline_artifact_response(
        artifact_type="slide_render",
        mime_type="image/png",
        bytes_data=PNG_BYTES,
        return_inline=True,
        response_mode="auto",
    )

    assert result["artifact"]["inline"] is True
    assert result["artifact"]["delivery"] == "base64_json"


def test_reads_payload_from_output_path(tmp_path):
    out = tmp_path / "slide.png"
    out.write_bytes(PNG_BYTES)

    result = build_inline_artifact_response(
        artifact_type="slide_render",
        mime_type="image/png",
        output_path=str(out),
        return_inline=True,
    )

    assert base64.b64decode(result["artifact"]["image_base64"]) == PNG_BYTES


def test_oversize_payload_returns_omission_without_failure():
    result = build_inline_artifact_response(
        artifact_type="slide_render",
        mime_type="image/png",
        bytes_data=b"abcdef",
        return_inline=True,
        max_inline_bytes=4,
    )

    artifact = result["artifact"]
    assert artifact["inline"] is False
    assert artifact["omitted_reason"] == "encoded image exceeds max_inline_bytes"
    assert artifact["encoded_bytes"] > artifact["max_inline_bytes"]
    assert "image_base64" not in artifact


def test_unsupported_mime_type_omits_payload():
    result = build_inline_artifact_response(
        artifact_type="document_render",
        mime_type="application/pdf",
        bytes_data=b"%PDF",
        return_inline=True,
    )

    assert result["artifact"]["inline"] is False
    assert result["artifact"]["omitted_reason"] == "unsupported mime type"
    assert "image_base64" not in result["artifact"]


def test_content_block_mode_is_explicitly_unsupported_for_now():
    result = build_inline_artifact_response(
        artifact_type="slide_render",
        mime_type="image/png",
        bytes_data=PNG_BYTES,
        return_inline=True,
        response_mode="content_block",
    )

    assert result["artifact"]["inline"] is False
    assert result["artifact"]["omitted_reason"] == "unsupported response mode"


def test_invalid_response_mode_rejected():
    with pytest.raises(ValueError, match="response_mode"):
        validate_inline_artifact_options(
            mime_type="image/png",
            response_mode="raw",
            encoded_or_raw_size=10,
            max_inline_bytes=100,
            allowed_mime_types={"image/png"},
        )


def test_metadata_passthrough_and_output_path_preserved(tmp_path):
    out = tmp_path / "sheet.png"
    metadata = {"sheet": "Sheet1", "range": "A1:H30"}

    result = build_inline_artifact_response(
        artifact_type="sheet_render",
        mime_type="image/png",
        bytes_data=PNG_BYTES,
        output_path=str(out),
        return_inline=True,
        metadata=metadata,
    )

    assert result["output_path"] == str(out)
    assert result["artifact"]["metadata"] == metadata


def test_empty_bytes_with_inline_returns_clear_omission():
    result = build_inline_artifact_response(
        artifact_type="slide_render",
        mime_type="image/png",
        bytes_data=b"",
        return_inline=True,
    )

    assert result["artifact"]["inline"] is False
    assert result["artifact"]["omitted_reason"] == "missing artifact bytes"
    assert result["artifact"]["encoded_bytes"] == 0


def test_base64_payload_not_logged(caplog):
    caplog.set_level(logging.INFO)
    result = build_inline_artifact_response(
        artifact_type="slide_render",
        mime_type="image/png",
        bytes_data=PNG_BYTES,
        return_inline=True,
    )

    assert result["artifact"]["image_base64"] not in caplog.text

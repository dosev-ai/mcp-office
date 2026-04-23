"""Runtime configuration for excelmcp metadata governance.

Pure stdlib module — no FastMCP imports, no COM imports.
Env vars are read at CALL TIME (not import time) so monkeypatch.setenv works in tests.
"""
from __future__ import annotations

import os
import warnings

REQUIRED_FIELDS: tuple[str, ...] = ("path", "workbook_name")

_VALID_POLICIES: frozenset[str] = frozenset({"strict", "lenient"})
_DEFAULT_POLICY: str = "strict"


def load_runtime_config() -> dict:
    """Read env vars, validate values, return config dict.

    Keys returned:
      metadata_policy: str             — "strict" | "lenient"
      metadata_contract_version: str   — CONTRACT_VERSION or override
      required_fields: tuple[str, ...] — REQUIRED_FIELDS constant
    """
    # 1. Read EXCEL_METADATA_POLICY, default "strict", validate
    raw_policy = os.getenv("EXCEL_METADATA_POLICY", _DEFAULT_POLICY).strip().lower()
    if raw_policy not in _VALID_POLICIES:
        raise ValueError(
            f"Invalid EXCEL_METADATA_POLICY={raw_policy!r}. "
            f"Must be one of: {sorted(_VALID_POLICIES)}"
        )

    # 2. Read EXCEL_METADATA_CONTRACT_VERSION, default from _metadata_contract
    from excelmcp._metadata_contract import CONTRACT_VERSION as _CV  # lazy import
    contract_version = os.getenv("EXCEL_METADATA_CONTRACT_VERSION", _CV).strip()

    return {
        "metadata_policy": raw_policy,
        "metadata_contract_version": contract_version,
        "required_fields": REQUIRED_FIELDS,
    }


def get_effective_policy() -> str:
    """Return the current metadata policy: "strict" or "lenient"."""
    raw = os.getenv("EXCEL_METADATA_POLICY", _DEFAULT_POLICY).strip().lower()
    if raw not in _VALID_POLICIES:
        warnings.warn(
            f"Invalid EXCEL_METADATA_POLICY={raw!r}, "
            f"falling back to {_DEFAULT_POLICY!r}. "
            f"Must be one of: {sorted(_VALID_POLICIES)}",
            UserWarning,
            stacklevel=2,
        )
        return _DEFAULT_POLICY   # safe fallback
    return raw


def is_strict() -> bool:
    """Return True when the effective policy is 'strict'."""
    return get_effective_policy() == "strict"


def get_contract_version() -> str:
    """Return the active contract version (env override or compiled-in default)."""
    from excelmcp._metadata_contract import CONTRACT_VERSION as _CV  # lazy import
    return os.getenv("EXCEL_METADATA_CONTRACT_VERSION", _CV).strip()


def get_session_mode() -> bool:
    """Return True when EXCEL_SESSION_MODE=true.

    Emits UserWarning if session mode is requested but EXCEL_ENABLE_COM is not
    'true', since the active-workbook bridge requires COM access.
    """
    enabled = os.getenv("EXCEL_SESSION_MODE", "false").strip().lower() == "true"
    if enabled and os.getenv("EXCEL_ENABLE_COM", "").strip().lower() != "true":
        warnings.warn(
            "EXCEL_SESSION_MODE=true but EXCEL_ENABLE_COM is not 'true' — "
            "session mode bridge will be unavailable; write tools will use openpyxl path.",
            UserWarning,
            stacklevel=2,
        )
    return enabled

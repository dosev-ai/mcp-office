"""mailmcp._core — Config, COM bootstrap, redaction, allowlist enforcement."""
from __future__ import annotations

import logging
import os
import re
import threading
import time
from dataclasses import dataclass, field, replace
from typing import Any

logger = logging.getLogger(__name__)

_DOMAIN_ALLOWLIST_RE = re.compile(
    r"^(?=.{1,253}\Z)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}\Z",
    re.IGNORECASE,
)


def _parse_nonnegative_int(raw: str, key: str) -> int:
    try:
        value = int(raw)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{key} must be a non-negative integer.") from exc
    if value < 0:
        raise ValueError(f"{key} must be a non-negative integer.")
    return value


def _parse_domain_list(raw: str, key: str = "OUTLOOK_ALLOWLIST_DOMAINS") -> list[str]:
    raw = raw.strip()
    if not raw:
        return []
    parts = [part.strip().lower() for part in raw.split(",")]
    if any(not part or not _DOMAIN_ALLOWLIST_RE.fullmatch(part) for part in parts):
        raise ValueError(
            f"{key} contains an invalid domain allowlist. Use comma-separated DNS domain names."
        )
    return list(dict.fromkeys(parts))


@dataclass
class OutlookConfig:
    allowlist_folders: list[str] = field(default_factory=lambda: ["Inbox", "Contacts"])
    max_items: int = 50
    max_body_chars: int = 4000
    attachment_max_mb: int = 10
    redact_mode: str = "none"
    enable_write: bool = False
    enable_send: bool = False
    enable_delete: bool = False
    enable_rules: bool = False
    allowlist_domains: list[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        for name in ("max_items", "max_body_chars", "attachment_max_mb"):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                raise ValueError(f"{name} must be a non-negative integer.")

    @classmethod
    def from_env(cls) -> "OutlookConfig":
        raw_folders = os.environ.get("OUTLOOK_ALLOWLIST_FOLDERS", "Inbox,Contacts")
        folders = [f.strip() for f in raw_folders.split(",") if f.strip()]
        return cls(
            allowlist_folders=folders,
            max_items=_parse_nonnegative_int(os.environ.get("OUTLOOK_MAX_ITEMS", "50"), "OUTLOOK_MAX_ITEMS"),
            max_body_chars=_parse_nonnegative_int(os.environ.get("OUTLOOK_MAX_BODY_CHARS", "4000"), "OUTLOOK_MAX_BODY_CHARS"),
            attachment_max_mb=_parse_nonnegative_int(os.environ.get("OUTLOOK_ATTACHMENT_MAX_MB", "10"), "OUTLOOK_ATTACHMENT_MAX_MB"),
            redact_mode=os.environ.get("OUTLOOK_REDACT_MODE", "none"),
            enable_write=os.environ.get("OUTLOOK_ENABLE_WRITE", "false").lower() == "true",
            enable_send=os.environ.get("OUTLOOK_ENABLE_SEND", "false").lower() == "true",
            enable_delete=os.environ.get("OUTLOOK_ENABLE_DELETE", "false").lower() == "true",
            enable_rules=os.environ.get("OUTLOOK_ENABLE_RULES", "false").lower() == "true",
            allowlist_domains=_parse_domain_list(
                os.environ.get("OUTLOOK_ALLOWLIST_DOMAINS", ""),
                "OUTLOOK_ALLOWLIST_DOMAINS",
            ),
        )


_REDACT_ORDER: list[str] = ["none", "emails", "emails+domains"]


@dataclass
class OutlookAccountOverride:
    """Per-account config restrictions parsed from OUTLOOK_ACCOUNT_N_* env vars."""

    email: str
    allowlist_folders: list[str] | None = None
    enable_write: bool | None = None
    enable_send: bool | None = None
    enable_delete: bool | None = None
    enable_rules: bool | None = None
    max_items: int | None = None
    max_body_chars: int | None = None
    redact_mode: str | None = None
    allowlist_domains: list[str] | None = None

    def __post_init__(self) -> None:
        for name in ("max_items", "max_body_chars"):
            value = getattr(self, name)
            if value is not None and (
                isinstance(value, bool) or not isinstance(value, int) or value < 0
            ):
                raise ValueError(f"{name} must be a non-negative integer when set.")


def _parse_folders_env(key: str) -> list[str] | None:
    raw = os.environ.get(key, "").strip()
    if not raw:
        return None
    return [f.strip() for f in raw.split(",") if f.strip()]


def _parse_bool_env(key: str) -> bool | None:
    raw = os.environ.get(key, "").strip().lower()
    if raw == "true":
        return True
    if raw == "false":
        return False
    return None


def _parse_int_env(key: str) -> int | None:
    raw = os.environ.get(key, "").strip()
    if not raw:
        return None
    return _parse_nonnegative_int(raw, key)


def _parse_str_env(key: str) -> str | None:
    raw = os.environ.get(key, "").strip()
    return raw if raw else None


def _parse_account_overrides() -> dict[str, OutlookAccountOverride]:
    overrides: dict[str, OutlookAccountOverride] = {}
    for n in range(1, 11):
        prefix = f"OUTLOOK_ACCOUNT_{n}_"
        email = os.environ.get(f"{prefix}EMAIL", "").strip().lower()
        if not email:
            continue
        domain_key = f"{prefix}ALLOWLIST_DOMAINS"
        overrides[email] = OutlookAccountOverride(
            email=email,
            allowlist_folders=_parse_folders_env(f"{prefix}ALLOWLIST_FOLDERS"),
            enable_write=_parse_bool_env(f"{prefix}ENABLE_WRITE"),
            enable_send=_parse_bool_env(f"{prefix}ENABLE_SEND"),
            enable_delete=_parse_bool_env(f"{prefix}ENABLE_DELETE"),
            enable_rules=_parse_bool_env(f"{prefix}ENABLE_RULES"),
            max_items=_parse_int_env(f"{prefix}MAX_ITEMS"),
            max_body_chars=_parse_int_env(f"{prefix}MAX_BODY_CHARS"),
            redact_mode=_parse_str_env(f"{prefix}REDACT_MODE"),
            allowlist_domains=_parse_domain_list(
                os.environ.get(domain_key, ""), domain_key
            ) or None,
        )
    return overrides


_config: OutlookConfig | None = None
_startup_config: OutlookConfig | None = None
_account_overrides: dict[str, OutlookAccountOverride] = {}
_startup_account_overrides: dict[str, OutlookAccountOverride] | None = None


def _load_account_overrides() -> None:
    global _account_overrides, _startup_account_overrides
    _account_overrides = _parse_account_overrides()
    if _startup_account_overrides is None:
        _startup_account_overrides = dict(_account_overrides)


_folder_obj_cache: dict[str, Any] = {}
_folder_cache_ts: float = 0.0
_folder_cache_lock: threading.Lock = threading.Lock()
_folder_cache_is_full: bool = False


def _get_folder_cache_ttl() -> float:
    raw = os.environ.get("OUTLOOK_FOLDER_CACHE_TTL_SECONDS", "300")
    try:
        val = float(raw)
    except (ValueError, TypeError):
        logger.warning("Invalid OUTLOOK_FOLDER_CACHE_TTL_SECONDS=%r, using 300", raw)
        return 300.0
    if val < 0:
        logger.warning("Negative OUTLOOK_FOLDER_CACHE_TTL_SECONDS=%r, clamping to 0", raw)
        return 0.0
    return val


def _is_folder_cache_valid() -> bool:
    with _folder_cache_lock:
        if not _folder_obj_cache:
            return False
        return (time.monotonic() - _folder_cache_ts) < _get_folder_cache_ttl()


def _invalidate_folder_cache() -> None:
    global _folder_cache_ts, _folder_cache_is_full
    with _folder_cache_lock:
        _folder_obj_cache.clear()
        _folder_cache_ts = 0.0
        _folder_cache_is_full = False


def _set_folder_cache(new_cache: dict[str, Any]) -> None:
    global _folder_cache_ts, _folder_cache_is_full
    with _folder_cache_lock:
        _folder_obj_cache.clear()
        _folder_obj_cache.update(new_cache)
        _folder_cache_ts = time.monotonic()
        _folder_cache_is_full = True


def _upsert_folder_cache_entry(key: str, folder: Any) -> None:
    global _folder_cache_ts, _folder_cache_is_full
    with _folder_cache_lock:
        _folder_obj_cache[key] = folder
        _folder_cache_ts = time.monotonic()
        _folder_cache_is_full = False


def _get_folder_cache_snapshot() -> dict | None:
    global _folder_cache_ts, _folder_cache_is_full
    with _folder_cache_lock:
        if not _folder_obj_cache:
            return None
        if (time.monotonic() - _folder_cache_ts) >= _get_folder_cache_ttl():
            return None
        snapshot = dict(_folder_obj_cache)
        ts_at_snapshot = _folder_cache_ts
    for _key, folder in snapshot.items():
        try:
            _ = folder.Name
        except Exception:
            with _folder_cache_lock:
                if _folder_cache_ts == ts_at_snapshot:
                    _folder_obj_cache.clear()
                    _folder_cache_ts = 0.0
                    _folder_cache_is_full = False
            return None
    return snapshot


def get_config() -> OutlookConfig:
    global _config, _startup_config
    if _config is None:
        _config = OutlookConfig.from_env()
        if _startup_config is None:
            _startup_config = _config
    return _config


def set_config(cfg: OutlookConfig) -> None:
    global _config
    _config = cfg


def get_startup_config() -> OutlookConfig:
    if _startup_config is None:
        get_config()
    return _startup_config  # type: ignore[return-value]


def _merge_config(g: OutlookConfig, ov: OutlookAccountOverride) -> OutlookConfig:
    if ov.allowlist_folders is not None:
        if "*" in g.allowlist_folders:
            merged_folders = ov.allowlist_folders
        else:
            global_lower = {f.lower() for f in g.allowlist_folders}
            merged_folders = [f for f in ov.allowlist_folders if f.lower() in global_lower]
    else:
        merged_folders = g.allowlist_folders

    if ov.allowlist_domains is not None:
        if not g.allowlist_domains:
            merged_domains = ov.allowlist_domains
        else:
            global_dom_lower = {d.lower() for d in g.allowlist_domains}
            merged_domains = [d for d in ov.allowlist_domains if d.lower() in global_dom_lower]
            if not merged_domains and ov.allowlist_domains:
                raise PermissionError("Global and account domain allowlists have no permitted intersection.")
    else:
        merged_domains = g.allowlist_domains

    def _bool_restrict(global_val: bool, override_val: bool | None) -> bool:
        if override_val is None:
            return global_val
        return global_val and override_val

    def _int_restrict(global_val: int, override_val: int | None) -> int:
        if override_val is None:
            return global_val
        return min(global_val, override_val)

    def _redact_restrict(global_val: str, override_val: str | None) -> str:
        if override_val is None:
            return global_val
        if global_val not in _REDACT_ORDER:
            logger.warning("unknown global redact_mode %r, treating as %r", global_val, _REDACT_ORDER[0])
        if override_val not in _REDACT_ORDER:
            logger.warning("unknown override redact_mode %r, treating as %r", override_val, _REDACT_ORDER[0])
        g_idx = _REDACT_ORDER.index(global_val) if global_val in _REDACT_ORDER else 0
        o_idx = _REDACT_ORDER.index(override_val) if override_val in _REDACT_ORDER else 0
        return _REDACT_ORDER[max(g_idx, o_idx)]

    return replace(
        g,
        allowlist_folders=merged_folders,
        allowlist_domains=merged_domains,
        enable_write=_bool_restrict(g.enable_write, ov.enable_write),
        enable_send=_bool_restrict(g.enable_send, ov.enable_send),
        enable_delete=_bool_restrict(g.enable_delete, ov.enable_delete),
        enable_rules=_bool_restrict(g.enable_rules, ov.enable_rules),
        max_items=_int_restrict(g.max_items, ov.max_items),
        max_body_chars=_int_restrict(g.max_body_chars, ov.max_body_chars),
        redact_mode=_redact_restrict(g.redact_mode, ov.redact_mode),
    )


def get_effective_config(account_email: str | None = None) -> OutlookConfig:
    g = get_config()
    if not account_email:
        return g
    ov = _account_overrides.get(account_email.strip().lower())
    if ov is None:
        return g
    return _merge_config(g, ov)


_EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
_DOMAIN_RE = re.compile(r"\b(?:[a-zA-Z0-9\-]+\.)+[a-zA-Z]{2,}\b")
_EMAIL_VALIDATE_RE = re.compile(r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}\Z')


def _redact(text: str, account_email: str | None = None) -> str:
    mode = get_effective_config(account_email).redact_mode
    if mode == "none" or not text:
        return text
    text = _EMAIL_RE.sub("[email]", text)
    if mode == "emails+domains":
        text = _DOMAIN_RE.sub("[domain]", text)
    return text


def _get_outlook() -> Any:
    try:
        import win32com.client  # type: ignore
        import pywintypes  # type: ignore
        try:
            return win32com.client.GetActiveObject("Outlook.Application")
        except (pywintypes.com_error, AttributeError):
            _invalidate_folder_cache()
            return win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:
        raise RuntimeError(
            "Cannot connect to Outlook COM. Ensure Outlook (classic desktop) is installed and not blocked by your org's programmatic-access policy."
        ) from exc


def _mapi():
    return _get_outlook().GetNamespace("MAPI")


_PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"


def _resolve_smtp_from_entry(entry: Any) -> str:
    try:
        smtp = entry.PropertyAccessor.GetProperty(_PR_SMTP_ADDRESS)
        if smtp and "@" in str(smtp):
            return str(smtp).strip().lower()
    except Exception:
        pass
    try:
        eu = entry.GetExchangeUser()
        if eu is not None:
            smtp2 = eu.PrimarySmtpAddress
            if smtp2 and "@" in str(smtp2):
                return str(smtp2).strip().lower()
    except Exception:
        pass
    addr = getattr(entry, "Address", None)
    if addr and "@" in str(addr):
        return str(addr).strip().lower()
    raise PermissionError("Cannot determine SMTP address for a recipient. Send blocked for safety.")


def _assert_domains_allowed(addresses: list[str], account_email: str | None = None) -> None:
    cfg = get_effective_config(account_email)
    if not cfg.allowlist_domains:
        return
    allowed = set(cfg.allowlist_domains)
    for addr in addresses:
        domain = addr.rsplit("@", 1)[-1].strip().lower()
        if domain not in allowed:
            raise PermissionError(f"Domain '{domain}' (from '{addr}') not in OUTLOOK_ALLOWLIST_DOMAINS.")


def _assert_allowed(folder_name: str, account_email: str | None = None) -> None:
    cfg = get_effective_config(account_email)
    allowed = [f.lower() for f in cfg.allowlist_folders]
    if folder_name.lower() not in allowed and "*" not in allowed:
        if account_email:
            raise PermissionError(
                f"Folder '{folder_name}' is not accessible for account '{account_email}'. Check OUTLOOK_ACCOUNT_N_ALLOWLIST_FOLDERS for that account."
            )
        raise PermissionError(
            f"Folder '{folder_name}' is not in the allowlist. Add it to OUTLOOK_ALLOWLIST_FOLDERS to enable access."
        )


def _assert_write_enabled(account_email: str | None = None) -> None:
    if not get_effective_config(account_email).enable_write:
        raise PermissionError(
            "Outlook write mutations are disabled. Set OUTLOOK_ENABLE_WRITE=true to enable draft, task, category, meeting-response, and other non-send mailbox mutations."
        )


def reload_config() -> dict:
    global _config, _account_overrides
    from mailmcp._config_policy import _assert_no_scope_expansion

    get_config()
    new_cfg = OutlookConfig.from_env()
    _new_acct_ovs = _parse_account_overrides()

    if _startup_config is not None:
        if new_cfg.enable_write and not _startup_config.enable_write:
            raise PermissionError(
                "reload_config: cannot enable 'enable_write' at runtime — the server was started with OUTLOOK_ENABLE_WRITE=false. Restart the server process with OUTLOOK_ENABLE_WRITE=true to allow non-send Outlook mutations."
            )
        if new_cfg.enable_send and not _startup_config.enable_send:
            raise PermissionError(
                "reload_config: cannot enable 'enable_send' at runtime — the server was started with OUTLOOK_ENABLE_SEND=false. Restart the server process with OUTLOOK_ENABLE_SEND=true to allow sending."
            )
        if new_cfg.enable_delete and not _startup_config.enable_delete:
            raise PermissionError(
                "reload_config: cannot enable 'enable_delete' at runtime — the server was started with OUTLOOK_ENABLE_DELETE=false. Restart the server process with OUTLOOK_ENABLE_DELETE=true to allow permanent deletion."
            )
        if new_cfg.enable_rules and not _startup_config.enable_rules:
            raise PermissionError(
                "reload_config: cannot enable 'enable_rules' at runtime — the server was started with OUTLOOK_ENABLE_RULES=false. Restart the server process with OUTLOOK_ENABLE_RULES=true to allow forwarding-rule mutations."
            )
        if _startup_config.allowlist_domains:
            startup_set = set(_startup_config.allowlist_domains)
            new_set = set(new_cfg.allowlist_domains)
            if not new_set or not new_set.issubset(startup_set):
                logger.warning(
                    "reload_config: allowlist_domains widening blocked — keeping startup value %s (attempted %s)",
                    _startup_config.allowlist_domains,
                    new_cfg.allowlist_domains,
                )
                new_cfg = replace(new_cfg, allowlist_domains=list(_startup_config.allowlist_domains))

        _assert_no_scope_expansion(_startup_config, new_cfg)
        if _startup_account_overrides is not None:
            _bool_flags = ("enable_write", "enable_send", "enable_delete", "enable_rules")
            for _email, _startup_ov in _startup_account_overrides.items():
                _startup_eff = _merge_config(_startup_config, _startup_ov)
                _new_ov = _new_acct_ovs.get(_email)
                _new_eff = _merge_config(new_cfg, _new_ov) if _new_ov is not None else new_cfg
                _assert_no_scope_expansion(_startup_eff, _new_eff)
                for _flag in _bool_flags:
                    if getattr(_new_eff, _flag) and not getattr(_startup_eff, _flag):
                        raise PermissionError(
                            f"reload_config: cannot expand '{_flag}' for account '{_email}' at runtime. The server was started with the flag disabled for this account. Restart the server process to change per-account security gates."
                        )

    _config = new_cfg
    _account_overrides = _new_acct_ovs
    _invalidate_folder_cache()
    logger.info(
        "reload_config applied: folders=%s max_items=%d write=%s send=%s delete=%s rules=%s",
        new_cfg.allowlist_folders,
        new_cfg.max_items,
        new_cfg.enable_write,
        new_cfg.enable_send,
        new_cfg.enable_delete,
        new_cfg.enable_rules,
    )
    return {
        "ok": True,
        "allowlist_folders": new_cfg.allowlist_folders,
        "max_items": new_cfg.max_items,
        "max_body_chars": new_cfg.max_body_chars,
        "attachment_max_mb": new_cfg.attachment_max_mb,
        "redact_mode": new_cfg.redact_mode,
        "enable_write": new_cfg.enable_write,
        "enable_send": new_cfg.enable_send,
        "enable_delete": new_cfg.enable_delete,
        "enable_rules": new_cfg.enable_rules,
        "allowlist_domains": new_cfg.allowlist_domains,
        "account_overrides": [{"email": ov.email} for ov in _account_overrides.values()],
    }


_load_account_overrides()

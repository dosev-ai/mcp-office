"""mcpshared — shared MCP contracts (ACP v1)."""
from mcpshared._acp_contract import (
    ACPAnnotation,
    ACPContent,
    ACPDeep,
    ACPDetail,
    ACPFocused,
    ACPIndex,
    _sanitize_text_field,
    validate_acp_annotations,
)

__all__ = [
    "ACPAnnotation",
    "ACPContent",
    "ACPDetail",
    "ACPIndex",
    "ACPFocused",
    "ACPDeep",
    "_sanitize_text_field",
    "validate_acp_annotations",
]

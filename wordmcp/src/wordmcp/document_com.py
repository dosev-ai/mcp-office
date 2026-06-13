"""document_com — thin shim for backward compatibility. Implementation is in wordmcp._com."""
from wordmcp._com._base import (  # noqa: F401
    _COM_AVAILABLE,
    _is_com_error,
    _com_err_msg,
    _check_com_enabled,
    _word_app_context,
    OfficeCOMError,
    ValidationError,
    NotAllowedError,
    _check_confirm,
    _check_path,
    _check_write,
)
from wordmcp._com._read import (  # noqa: F401
    _CTRL_RE,
    _sanitize_revision_text,
    list_tracked_changes,
)
from wordmcp._com._write import (  # noqa: F401
    accept_all_track_changes,
    reject_all_track_changes,
    manage_tracked_changes,
)
from wordmcp._com._export import (  # noqa: F401
    _check_output_path,
    _check_output_path_html,
    export_as_pdf,
    export_document,
)

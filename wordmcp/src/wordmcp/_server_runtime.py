from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Any

from fastmcp.exceptions import ToolError
from wordmcp.document_docx import FileError as _FileError

logger = logging.getLogger(__name__)


class ServerRuntime:
    def __init__(
        self,
        *,
        get_doc: Callable[[], Any],
        get_word_error: Callable[[], type[Exception]],
        get_com: Callable[[], Any],
        get_com_loaded: Callable[[], bool],
        require_com: Callable[[], None],
    ) -> None:
        self._get_doc = get_doc
        self._get_word_error = get_word_error
        self._get_com = get_com
        self._get_com_loaded = get_com_loaded
        self._require_com = require_com

    def doc(self) -> Any:
        return self._get_doc()

    def com(self) -> Any:
        return self._get_com()

    def word_error(self) -> type[Exception]:
        return self._get_word_error()

    def is_com_loaded(self) -> bool:
        return bool(self._get_com_loaded())

    def require_com(self) -> None:
        self._require_com()

    def call_doc(self, fn_name: str, *args: Any, **kwargs: Any) -> Any:
        return self._call(self.doc(), fn_name, *args, **kwargs)

    def call_com(self, fn_name: str, *args: Any, **kwargs: Any) -> Any:
        self.require_com()
        return self._call(self.com(), fn_name, *args, **kwargs)

    def _call(self, backend: Any, fn_name: str, *args: Any, **kwargs: Any) -> Any:
        try:
            return getattr(backend, fn_name)(*args, **kwargs)
        except self.word_error() as exc:
            logger.error("WordMCPError in tool call '%s': %s", fn_name, exc)
            # W2-004: FileError messages may contain full file-system paths—sanitise.
            # All other WordMCPError subtypes carry intentional, user-facing
            # messages (validation, confirm gate, etc.) and are safe to expose.
            if isinstance(exc, _FileError):
                raise ToolError(f"Operation failed ({type(exc).__name__})") from None
            raise ToolError(str(exc)) from None
        except Exception as exc:
            raise ToolError(f"Internal error ({type(exc).__name__})") from None

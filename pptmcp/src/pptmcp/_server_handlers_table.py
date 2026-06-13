"""FastMCP handlers for table mutation tools.

add_column      — append column to a table shape.
edit_table_cell — write text to a cell by row/col index.
set_column_width — resize a column.

No COM logic. Delegates all python-pptx work to _pptx_table_mutations.
"""
from __future__ import annotations

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

from pptmcp._pptx_table_mutations import (
    add_column,
    edit_table_cell,
    set_column_width,
)
from pptmcp.presentation_pptx import PPTMCPError, NotAllowedError, ValidationError


def register_table_tools(mcp: FastMCP) -> None:
    """Register add_column, edit_table_cell, set_column_width on the FastMCP instance."""

    @mcp.tool()
    def pptmcp_add_column(
        path: str,
        slide_index: int,
        shape_id: int,
        width_inches: float = 1.0,
        confirm: bool = False,
    ) -> dict:
        """Append a new column to an existing table shape on a slide.

        Copies the last column's cell formatting into the new cells (text cleared).
        Requires PPT_ENABLE_WRITE=true AND confirm=True.

        Args:
            path: Path to the .pptx file.
            slide_index: Zero-based slide index.
            shape_id: shape_id of the table shape (use list_shapes to find it).
            width_inches: Width of the new column in inches (default 1.0).
            confirm: Must be True to proceed (safety gate).

        Returns:
            {status, slide_index, shape_id, column_count, width_inches}
        """
        try:
            return add_column(
                path=path,
                slide_index=slide_index,
                shape_id=shape_id,
                width_inches=width_inches,
                confirm=confirm,
            )
        except (NotAllowedError, ValidationError, PPTMCPError) as exc:
            raise ToolError(str(exc)) from exc
        except ValueError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def pptmcp_edit_table_cell(
        path: str,
        slide_index: int,
        shape_id: int,
        row: int,
        col: int,
        text: str,
        confirm: bool = False,
    ) -> dict:
        """Write text to a table cell identified by zero-based row and column index.

        Only the target cell is changed; all other cells remain unchanged.
        Requires PPT_ENABLE_WRITE=true AND confirm=True.

        Args:
            path: Path to the .pptx file.
            slide_index: Zero-based slide index.
            shape_id: shape_id of the table shape.
            row: Zero-based row index.
            col: Zero-based column index.
            text: Text to write into the cell.
            confirm: Must be True to proceed (safety gate).

        Returns:
            {status, slide_index, shape_id, row, col, text}
        """
        try:
            return edit_table_cell(
                path=path,
                slide_index=slide_index,
                shape_id=shape_id,
                row=row,
                col=col,
                text=text,
                confirm=confirm,
            )
        except (NotAllowedError, ValidationError, PPTMCPError) as exc:
            raise ToolError(str(exc)) from exc
        except ValueError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def pptmcp_set_column_width(
        path: str,
        slide_index: int,
        shape_id: int,
        col: int,
        width_inches: float,
        confirm: bool = False,
    ) -> dict:
        """Set the width of a table column in inches.

        Requires PPT_ENABLE_WRITE=true AND confirm=True.

        Args:
            path: Path to the .pptx file.
            slide_index: Zero-based slide index.
            shape_id: shape_id of the table shape.
            col: Zero-based column index.
            width_inches: New column width in inches.
            confirm: Must be True to proceed (safety gate).

        Returns:
            {status, slide_index, shape_id, col, width_inches}
        """
        try:
            return set_column_width(
                path=path,
                slide_index=slide_index,
                shape_id=shape_id,
                col=col,
                width_inches=width_inches,
                confirm=confirm,
            )
        except (NotAllowedError, ValidationError, PPTMCPError) as exc:
            raise ToolError(str(exc)) from exc
        except ValueError as exc:
            raise ToolError(str(exc)) from exc

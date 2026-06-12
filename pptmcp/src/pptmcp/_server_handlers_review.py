"""
Review / contract / evidence MCP tool registrations for pptmcp.
"""
from __future__ import annotations

__all__: list[str] = []

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

import pptmcp.review_pptx as review
from pptmcp import contract_pptx as contract
from pptmcp import presentation_pptx as prs
from pptmcp.presentation_pptx import PPTMCPError


def register_review_tools(mcp: FastMCP) -> None:
    """Register all 4 review/contract/evidence tools on the given FastMCP instance."""

    @mcp.tool()
    def declare_slide_contract(
        slide_path: str,
        spec: dict,
        confirm: bool = False,
    ) -> dict:
        """Persist a slide-export contract beside the PPTX (RCI-US-01).

        Writes ``contract.json`` in the same directory as the PPTX and
        returns ``{contract_path, checksum_sha256, contract_version}``.
        The spec is validated against the v1.0 schema before the file is
        committed; failures leave no partial file on disk.

        Write-gated (PPT_ENABLE_WRITE) + confirm=True + slide_path allowlisted.
        """
        try:
            return contract.declare_slide_contract(slide_path, spec, confirm=confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def validate_contract(contract_path: str) -> dict:
        """Validate a JSON Output Contract file against the pptmcp v1.0 schema.

        Returns a dict with keys:
          valid (bool), contract_version (str), slide_count (int), errors (list[str]).

        Raises a ToolError for path/IO failures (allowlist denial, file not found,
        bad extension, file too large, malformed JSON, missing contract_version,
        boolean in a numeric field).
        Schema mismatches return valid=False with a populated errors list.
        """
        try:
            return contract.validate_contract(contract_path)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def check_presentation_against_contract(path: str, contract_path: str) -> dict:
        """Check a presentation against its Output Contract specification.

        Returns a dict with keys:
          passed (bool), slide_count (int), findings (list[dict]).

        Each finding has: slide_index (int or null), criterion (str), detail (str),
        and optionally severity='warning' for non-blocking issues.
        A passing result returns {"passed": true, "findings": []}.

        Raises a ToolError for any path/IO access error on either file.
        """
        try:
            return contract.check_presentation_against_contract(path, contract_path)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def review_slide_export(
        path: str,
        slide_index: int,
        png_path: str,
        contract_path: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """Review a slide export against quality criteria.

        Checks TEXT_OVERFLOW (chars/sq.in heuristic), SHAPE_OUTSIDE_SLIDE (bounding box vs
        slide dimensions), and TABLE_CONTENT_CLIPPING (table row heights vs. shape bounding box)
        for the specified slide. Write-gated (PPT_ENABLE_WRITE) + confirm=True required to
        write the audit artefact.

        Args:
            path: Path to the .pptx file (must be in PPT_ALLOWLIST_ROOTS).
            slide_index: 0-based slide index.
            png_path: Path to the exported PNG (must exist and be in PPT_ALLOWLIST_ROOTS).
            contract_path: Optional path to Output Contract JSON for reference.
            confirm: Must be True (with PPT_ENABLE_WRITE=true) to write review_result.json.

        Returns:
            dict with keys: slide_index, png_path, pptx_path, passed, findings,
                            checks_run, findings_count.
        """
        try:
            return review.review_slide_export(
                path=path,
                slide_index=slide_index,
                png_path=png_path,
                contract_path=contract_path,
                confirm=confirm,
            )
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def produce_evidence_bundle(
        path: str,
        review_results: list,
        contract_path: str | None = None,
        exports_dir: str | None = None,
    ) -> dict:
        """Aggregate review_slide_export results into a machine-verifiable evidence bundle.

        review_results: list of dicts from review_slide_export().
        Returns: evidence bundle dict with bundle_version, overall_passed, slides, total_findings.
        """
        try:
            return prs.produce_evidence_bundle(path, review_results, contract_path, exports_dir)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

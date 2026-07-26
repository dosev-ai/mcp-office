# Governed Excel-to-PowerPoint workflow — public proof fixture

This folder provides a public-safe, synthetic documentation fixture for a bounded cross-application Office workflow:

```text
structured table -> deterministic deck specification -> PowerPoint summary and detail slides
```

It is intended to make the proof inspectable without exposing employer, client, supplier, personal, or machine-specific data.

## Contents

- [`synthetic_supplier_performance.csv`](synthetic_supplier_performance.csv) — synthetic finance/procurement-style source table.
- [`deck_spec.json`](deck_spec.json) — deterministic presentation contract derived from the source shape.
- [`expected_output.md`](expected_output.md) — expected slide views, governed visual preview, and explicit non-claims.

## Delivered proof boundary

The underlying proof established the following bounded path:

1. accept structured Excel or typed tabular input;
2. normalize the source into an inspectable `deck_spec`;
3. generate title, executive-summary/KPI, and detail-table slides;
4. enforce approved-path and explicit-write controls;
5. preserve lineage between input, specification, output, and validation evidence.

This documentation bundle exposes the public-safe input and specification plus an explanatory output preview. It does **not** present a downloadable deck as if it were generated during this publication step.

## Control model

The fixture describes the same control principles used across MCP Office:

- operations remain inside approved file roots;
- mutation requires explicit runtime and operation-level permission;
- data is normalized before bounded slide operations are selected;
- input, specification, output, and validation remain attributable;
- synthetic data is used for public documentation.

## Deliberate limitations

This proof does not claim:

- Gantt or timeline-slide generation;
- automated visual render-and-repair;
- MailMCP distribution;
- unrestricted Python or shell execution;
- unattended approval or publication;
- public release of WorkflowRuntime.

## Public project

MCP Office currently publishes `excelmcp`, `pptmcp`, and `wordmcp`. `mailmcp` is coming next after its package-specific release gates pass.

Return to the [MCP Office repository](../../../README.md).

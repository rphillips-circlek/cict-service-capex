# CapEx workflow Azure readiness

## Summary

| Component | Build | Completeness | Deployability |
|---|---|---|---|
| `capex-flow` | PASS | PASS | FAIL |

Detected stack: Python 3.9+ CLI/batch workflow using Microsoft Agent Framework and Autarch.

The CapEx implementation is ready for a controlled demo against Azure OpenAI using a dedicated finite batch container. Durable state and authenticated human review remain production hardening work.

## Findings

### Warnings

- **W-LOCAL-STATE:** Case snapshots, review events, dashboards, decision packages, and audit exports use a local workspace. Use persistent storage for a hosted demo; the target architecture should use Azure SQL for work/decision state and immutable Blob Storage for evidence.
- **W-HEALTH:** The CapEx process is finite batch work with no HTTP health endpoint. A Container Apps Job fits the demo better than an always-on web app.
- **W-REVIEW-AUTH:** Human reviews record a supplied name but do not authenticate that reviewer. This is acceptable only for a controlled demo and must be replaced by Entra-authenticated identity for production.

## Positive readiness evidence

- The root package manifest defines the Microsoft Agent Framework, Azure Identity, Pydantic, and cryptography dependencies needed by the current CapEx flow.
- The workflow has an explicit live MAF mode with no silent rules fallback.
- The workflow includes bounded model calls, structured outputs, deterministic checks, signed why-records, review binding, and generated decision/audit artifacts.
- Sample ABBYY/Vantage invoice exports are present for a repeatable demo.
- Focused tests exist for the CapEx flow, though no test or build command was executed during this static prereq phase.

## Post-Fix Re-evaluation

The Azure-hosting blocker is resolved:

- Added `Dockerfile.capex`, a non-root Python 3.12 image that installs the `capex` dependency set, packages the workflow and sample ABBYY inputs, and persists its workspace beneath `/data`.
- Added `.dockerignore`.
- Preserved the original generic Autarch `Dockerfile` unchanged.

Re-evaluated verdicts: Build PASS, Completeness PASS, Deployability PASS. Overall health is ready with caveats because cloud state durability and reviewer authentication are not yet production-grade.

## Latest-scope alignment

For the next demo, use `examples/capex_flow.py`, not the older end-to-end invoice extraction/tax example. The latest architecture consumes one upstream extraction and excludes invoice re-extraction, taxability determination, automatic financial-system writeback, and project-close bundling from Phase 1.

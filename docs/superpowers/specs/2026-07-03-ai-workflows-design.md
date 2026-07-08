# AI Workflows Documentation Design

Date: 2026-07-03

## Goal

Make the two recurring workstreams readable from GitHub and usable by other AI models:

- Updating the TREE workbook.
- Producing macro daily conclusions.

The output should turn existing work logs and rule files into stable project documentation without changing the Excel workbooks or workflow scripts.

## Recommended Structure

Add a documentation entry point:

- `docs/AI_WORKFLOWS.md`

Add one folder for AI-readable workflow manuals:

- `docs/ai_workflows/TREE_UPDATE_WORKFLOW.md`
- `docs/ai_workflows/MACRO_CONCLUSION_WORKFLOW.md`

Update existing AI onboarding documents:

- `docs/project_info/AI_CONTEXT.md`
- `docs/project_info/AGENTS.md`

## Content Design

`docs/AI_WORKFLOWS.md` should be the shortest entry point. It will tell an AI which workflow to choose, which files are authoritative, and which documents to read first.

`TREE_UPDATE_WORKFLOW.md` should summarize the operational flow already captured in `tools/tree_daily_update/TREE_UPDATE_RULES.md` and `logs/WORK_LOG.md`: identify latest inputs, verify the daily workbook, preserve TREE layout, sync data, update write time and trend text, refresh TREE chart caches, audit XML/visual output, and record delivery facts.

`MACRO_CONCLUSION_WORKFLOW.md` should turn `logs/宏观日报结论输出工作日志_20260703.md` into a stable runbook: use latest TREE and daily bottom data, rebuild the auxiliary judgement table, run first-round indicator screening, make face-level conclusions, check reverse evidence, do second-round PPT indicator selection, produce standard conclusion sentences, and update TREE/PPT only after conclusions are confirmed.

`AI_CONTEXT.md` and `AGENTS.md` should link to the new workflow entry point so external AI models find the new manuals early.

## Boundaries

- Do not edit any `.xlsx` or `.pptx` files.
- Do not move existing logs.
- Keep the existing detailed work logs as history; the new docs are distilled operating manuals.
- Preserve `tools/tree_daily_update/TREE_UPDATE_RULES.md` as the authority for calculation, formatting, chart, and audit rules.
- Preserve `logs/宏观日报结论输出工作日志_20260703.md` as the source work log for conclusion-output reasoning.

## Verification

Verify that:

- The new Markdown files exist.
- The new docs link to existing authoritative files.
- `AI_CONTEXT.md` and `AGENTS.md` mention `docs/AI_WORKFLOWS.md`.
- The final change is documentation-only.

## Self-Review

- No placeholders remain.
- The structure is focused on the two approved workflows.
- The design does not require network access or GitHub changes from the local environment.
- The only limitation is that `git` is unavailable in the current shell, so the design document cannot be committed from this environment.

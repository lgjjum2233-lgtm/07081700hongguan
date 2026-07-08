# AI Workflows Documentation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add GitHub-readable AI workflow documentation for TREE updates and macro daily conclusion output.

**Architecture:** Add one workflow index under `docs/`, two focused workflow manuals under `docs/ai_workflows/`, and links from the existing AI onboarding files. Keep logs and rule files as authoritative history while the new docs serve as distilled runbooks.

**Tech Stack:** Markdown documentation, PowerShell verification commands, existing project files.

---

## File Structure

- Create `docs/AI_WORKFLOWS.md`: top-level entry point for AI models.
- Create `docs/ai_workflows/TREE_UPDATE_WORKFLOW.md`: distilled TREE update runbook.
- Create `docs/ai_workflows/MACRO_CONCLUSION_WORKFLOW.md`: distilled macro conclusion output runbook.
- Modify `docs/project_info/AI_CONTEXT.md`: add the new workflow entry to the AI reading order and workflow notes.
- Modify `docs/project_info/AGENTS.md`: add the new workflow entry to agent rules.
- Verify with `Test-Path`, `rg`, and Markdown link spot checks.

### Task 1: Add Workflow Index

**Files:**
- Create: `D:\AI\宏观框架分析\docs\AI_WORKFLOWS.md`

- [ ] **Step 1: Create the workflow index**

Write a Markdown file that:

- Explains that this directory is for AI-readable recurring workflows.
- Links to `docs/ai_workflows/TREE_UPDATE_WORKFLOW.md`.
- Links to `docs/ai_workflows/MACRO_CONCLUSION_WORKFLOW.md`.
- Lists the authoritative source files: `tools/tree_daily_update/TREE_UPDATE_RULES.md`, `logs/WORK_LOG.md`, and `logs/宏观日报结论输出工作日志_20260703.md`.

- [ ] **Step 2: Verify the index exists**

Run:

```powershell
Test-Path 'D:\AI\宏观框架分析\docs\AI_WORKFLOWS.md'
```

Expected: `True`

### Task 2: Add TREE Update Workflow Manual

**Files:**
- Create: `D:\AI\宏观框架分析\docs\ai_workflows\TREE_UPDATE_WORKFLOW.md`

- [ ] **Step 1: Create the TREE runbook**

Write sections for:

- Purpose.
- Authoritative sources.
- Required inputs.
- Standard workflow.
- Data and calculation rules.
- Trend chart rules.
- Audit checklist.
- Delivery summary requirements.
- Common mistakes to avoid.

- [ ] **Step 2: Verify required TREE terms**

Run:

```powershell
rg -n "写入时间|趋势图|TREE_UPDATE_RULES|交付" 'D:\AI\宏观框架分析\docs\ai_workflows\TREE_UPDATE_WORKFLOW.md'
```

Expected: all four terms appear at least once.

### Task 3: Add Macro Conclusion Workflow Manual

**Files:**
- Create: `D:\AI\宏观框架分析\docs\ai_workflows\MACRO_CONCLUSION_WORKFLOW.md`

- [ ] **Step 1: Create the conclusion runbook**

Write sections for:

- Purpose.
- Authoritative sources.
- Data source rules.
- Auxiliary judgement table role.
- First-round screening.
- Face-level conclusion synthesis.
- Reverse evidence check.
- Second-round PPT selection.
- Standard sentence format.
- PPT update rules.
- Delivery summary requirements.
- Common mistakes to avoid.

- [ ] **Step 2: Verify required conclusion terms**

Run:

```powershell
rg -n "辅助判断表|第一轮|第二轮|反向证据|PPT" 'D:\AI\宏观框架分析\docs\ai_workflows\MACRO_CONCLUSION_WORKFLOW.md'
```

Expected: all five terms appear at least once.

### Task 4: Link Workflows from AI Onboarding

**Files:**
- Modify: `D:\AI\宏观框架分析\docs\project_info\AI_CONTEXT.md`
- Modify: `D:\AI\宏观框架分析\docs\project_info\AGENTS.md`

- [ ] **Step 1: Update AI_CONTEXT**

Add `docs/AI_WORKFLOWS.md` to the priority reading order and mention that TREE updates and conclusion output should follow the new workflow manuals.

- [ ] **Step 2: Update AGENTS**

Add `docs/AI_WORKFLOWS.md` to the priority reading order and add a short rule that recurring work must use the relevant workflow manual before acting.

- [ ] **Step 3: Verify onboarding links**

Run:

```powershell
rg -n "AI_WORKFLOWS|TREE_UPDATE_WORKFLOW|MACRO_CONCLUSION_WORKFLOW" 'D:\AI\宏观框架分析\docs\project_info\AI_CONTEXT.md' 'D:\AI\宏观框架分析\docs\project_info\AGENTS.md'
```

Expected: `AI_WORKFLOWS` appears in both files.

### Task 5: Final Documentation Verification

**Files:**
- Inspect all created and modified Markdown files.

- [ ] **Step 1: Verify file set**

Run:

```powershell
Test-Path 'D:\AI\宏观框架分析\docs\AI_WORKFLOWS.md'
Test-Path 'D:\AI\宏观框架分析\docs\ai_workflows\TREE_UPDATE_WORKFLOW.md'
Test-Path 'D:\AI\宏观框架分析\docs\ai_workflows\MACRO_CONCLUSION_WORKFLOW.md'
```

Expected: three `True` lines.

- [ ] **Step 2: Verify this is documentation-only**

Run:

```powershell
rg -n "AI_WORKFLOWS|TREE_UPDATE_WORKFLOW|MACRO_CONCLUSION_WORKFLOW|宏观日报结论输出工作日志" 'D:\AI\宏观框架分析\docs'
```

Expected: matches only in Markdown documentation files.

- [ ] **Step 3: Record git limitation**

Run:

```powershell
Get-Command git -ErrorAction SilentlyContinue
```

Expected in current environment: no output, because `git` is not available in this shell.

## Self-Review

- Spec coverage: all design requirements map to Tasks 1-5.
- Placeholder scan: no TODO, TBD, or deferred implementation language remains.
- Type consistency: file names and links match the design spec exactly.
- Environment note: commit steps are omitted because the current shell cannot locate `git`.

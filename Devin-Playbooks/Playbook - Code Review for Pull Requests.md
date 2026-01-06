# Playbook: Code Review for Pull Requests — Devin Playbook

---

## Overview

Deterministic workflow for reviewing a GitHub PR. Prioritizes correctness, security, reliability, performance, and maintainability; produces **clear line-anchored comments** and **one overall decision**; and ends with a lightweight **documentation PR** capturing the review summary. Devin does **not** modify code, run tests, or merge.

---

## 1. Background Data, Documents, and Images

- PR URL/number; repository read access.
- PR description, linked issues, acceptance criteria, screenshots (if any).
- Project standards: coding, contribution, security/perf, test strategy; architecture docs/ADRs; API contracts.
- Tooling: GitHub CLI (`gh`), local git (read-only clone).
- CI status for the PR (checks, artifacts).

## 2. Detailed Task Description & Rules

**Objective:** Review the specified PR and submit inline comments plus an overall decision. Do **not** modify code, run tests, or merge.

**Rules**

- **Scope:** Review PR changes and necessary context; avoid unrelated refactors.
- **Focus order:** Correctness & security → reliability/error handling → performance → readability/maintainability → consistency/style.
- **Evidence-based:** Reference files/lines, explain **why**, propose concrete alternatives.
- **Tone:** Professional, specific, and actionable; ask when intent is unclear.
- **Generated/Vendored files:** Sanity-check only (lockfiles, `dist/`, `.min.js`, images).
- **No execution:** Don’t run code or tests unless explicitly instructed.

**Assumptions**

- `gh` is authenticated and can fetch PR metadata/diffs.
- Read-only checkout is permitted.

## 3. Examples

**Inline comment (specific, constructive)**

```
**[file: src/user/service.ts | line 143]****[major]** Potential null dereference: `user.profile` may be undefined when `getProfile()` returns null.**Why**: This path is not guarded by the earlier return.**Suggestion**:```tsconst profile = await getProfile(id);if (!profile) {  return err(new NotFoundError("Profile not found"));}
```

````

**GitHub “suggested change”**
```diff
- const items = await db.findAll({ limit: 1000 });
+ // Avoid loading large sets; paginate or stream
+ const items = await db.findAll({ limit: 100, offset });

````

**Overall review (body)**

```
## Summary- Scope: Add optimistic concurrency control to Order updates.- Risk: Medium (touches write path).- Positives: Clear separation of concerns; good test names.- Concerns:  1) Race window remains in `updateIfVersionMatches` (see inline).  2) Missing negative test for stale version.  3) Exposes internal error messages to clients.## DecisionRequest changes.## Next steps1) Add negative tests.2) Mask DB errors at API boundary.3) Move version check inside the transaction.
```

## 4. Conversation History

Use PR discussion only to clarify intent. Summarize key clarifications in the overall review body.

## 5. Immediate Task Description or Request

Review the given PR URL/number and provide inline comments plus one overall decision.

## 6. Thinking Step by Step / Deep Breath (Reasoning Checklist)

**Triage**

- Title/description adequate? Linked issues match scope? Draft vs ready? CI status?

**Correctness**

- Data flow and edge cases (null/undefined, off-by-one, boundaries, timezone/locale).
- Concurrency/transactions; idempotency; unexpected mutation/shared state.

**Security**

- Input validation; AuthN/AuthZ; injection risks; safe crypto APIs.
- Secrets/logging: no credentials or sensitive data in code or logs.

**Reliability & Error Handling**

- Fail fast on invariants; no swallowed errors; actionable messages (no internals to clients).

**Performance**

- N+1 queries; O(n²) loops; synchronous blocking in hot paths; redundant work.

**Readability/Maintainability**

- Clear naming; cohesive functions; comments explain **why**; no dead code; reasonable diff size.

**Consistency**

- Matches project structure, conventions, lint/format.

**Testing**

- Coverage of happy/failure paths; deterministic fixtures; mocks/stubs appropriate.

**Docs/Changelog**

- README/API docs/migrations updated if behavior changes.

## 7. Output Formatting

**Deliverables**

1. **Inline review comments** anchored to files/lines.
2. **One overall decision**: **Approve**, **Request changes**, or **Comment**.

**Severity tags (prefix each comment)**

- **[blocker]** must change (correctness/security)
- **[major]** strongly recommended (reliability/perf/maintainability)
- **[minor]** nice-to-have
- **[nit]** non-blocking style/readability

**Inline comment template**

```
**[file: <path> | line <n>]****[<blocker|major|minor|nit>]** <issue/observation in 1–3 sentences>**Why**: <impact/rationale>**Suggestion**:```<language-or-text><short code or concrete action>
```

````

**Overall review template**
```md
## Summary
- Scope: <one-liner>
- Risk: <Low|Medium|High> (why)
- Positives: <bullets>
- Concerns: <bullets pointing to inlines>
## Decision
<Approve | Request changes | Comment>
## Next steps
<ordered list>

````

## 8. Prefilled Response (if any)

N/A

---

## User Inputs

- **PR URL / number** (required)
- **Areas of focus** (optional)
- **Links to project standards/guidelines** (optional)

## Procedure

### 1) Prepare

```
gh auth statusgh pr view <PR> --json title,number,author,baseRefName,headRefName,changedFiles,additions,deletions,isDraft,mergeStateStatus,labels,body
```

- Read title/description; note base/head; scan linked issues and CI status.

### 2) Check out PR (read-only)

```
gh pr checkout <PR>git branch --show-current   # head branch
```

### 3) Analyze diffs efficiently

```
gh pr diff <PR> --name-onlygh pr diff <PR> --patch# For a single commit or file:gh pr diff <PR> --commit <SHA>gh pr diff <PR> --path <file>
```

- Down-prioritize generated/binary files (`package-lock.json`, `dist/`, `.min.js`, images).

### 4) Systematic review (use the checklist in §6)

- Work file-by-file or commit-by-commit.
- Open surrounding context for cross-cutting logic.

### 5) Draft and post inline comments

```
gh pr review <PR> \  --comment \  --body "See line 143: **[major]** Possible null deref; guard \`profile\` before use." \  --path src/user/service.ts --line 143# Tip: escape backticks with \`
```

- Prefer **suggested changes** in the UI when the fix is obvious.

### 6) Submit one overall decision

```
# Choose exactly one:gh pr review <PR> --approve --body "Looks good; minor nits noted."gh pr review <PR> --request-changes --body "$(cat review_summary.md)"gh pr review <PR> --comment --body "$(cat review_summary.md)"
```

### 7) PR-First Documentation (required)

Create a small **review log PR** summarizing your review (no product code changes).

```
git switch -c devin/pr-review-<PRNUM>mkdir -p docs/pr-reviewscat > docs/pr-reviews/<PRNUM>.md <<'EOF'# PR Review Log <PRNUM><paste your overall review body here>EOFgit add docs/pr-reviews/<PRNUM>.mdgit commit -m "docs(review): log review for PR #<PRNUM>"git push -u origin devin/pr-review-<PRNUM>gh pr create --fill --title "Docs: Review log for PR #<PRNUM>" --body "Adds review summary to docs/pr-reviews/<PRNUM>.md"
```

> If the repo forbids docs changes, open the PR as draft; the branch still satisfies PR-first workflow and can be closed after sign-off.

## Output Specification

- Inline comments with severity and concrete suggestions where possible.
- One overall decision with clear summary and next steps.
- A doc PR (`devin/pr-review-<PRNUM>`) adding `docs/pr-reviews/<PRNUM>.md` mirroring the overall review body.

## Forbidden Actions

- Don’t push code changes to the PR branch.
- Don’t merge the PR.
- Don’t run tests or the app unless explicitly asked.
- Don’t expose secrets or paste large proprietary code blocks.
- Don’t approve if critical issues remain.
- Don’t demand large refactors outside scope without justification.

## Advice and Pointers

- Lead with correctness & security; batch nits; keep comments focused.
- Use consistent severity tags and line anchors; call out what’s well done.
- Reference project style guides when proposing style changes.
- Prefer fixes that align with existing architecture.

## PR & Branching

- **Doc branch**: `devin/pr-review-<PRNUM>`
- **Commit style**: concise, imperative (e.g., `docs(review): log review for PR #123`)
- **PR body**: summary, areas reviewed, notable risks, follow-ups (if any)

## Completion Criteria

- Inline comments posted and anchored to relevant lines/files.
- Overall decision submitted (**Approve**, **Request changes**, or **Comment**).
- Project guidelines (if provided) are applied and referenced.
- Summary posted even when approving.
- Documentation PR opened from a `devin/` branch with the review log.
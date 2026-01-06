# Playbook: PR Documentation

## Overview

This playbook defines a deterministic workflow to produce high‑quality Pull Request (PR) documentation across GitHub, GitLab, and Bitbucket. It standardizes branch naming, commit hygiene, PR templates, checklists, artifacts (screenshots, GIFs, logs), cross‑links to issues/epics, CHANGELOG/Release Notes updates, and reviewer/label setup. The outcome is a well‑scoped PR with a complete, compliant description, validation evidence, and a clean path to merge.

## 1. Background Data, Documents, and Images

- Repository URL and default branch (e.g., `main`)
- Target tracker items: issue/ticket/epic IDs
- Contribution guidelines (`CONTRIBUTING.md`), `CODEOWNERS`, security guidelines
- Release/versioning policy (SemVer), Conventional Commits policy (if any)
- Existing PR/MR templates and label taxonomy
- CI/CD status checks required for merge
- CHANGELOG/RELEASE_NOTES location and format
- ADRs/design docs/acceptance criteria relevant to the change
- Evidence assets: screenshots/GIFs, benchmark results, logs (scrubbed), diagrams
- Test artifacts: unit/integration/e2e results, coverage reports

## 2. Detailed Task Description & Rules

**Primary Objective**: Create and submit a reviewable PR with complete, accurate, and self‑contained documentation and validation evidence.

**Core Rules**

- **Scope**: One logical change per PR. If scope drifts, split into separate PRs.
- **Branch**: Use `devin/pr-docs-<short-slug>` (e.g., `devin/pr-docs-improve-search`) from the latest target branch.
- **Commits**: Clean history. Prefer Conventional Commits (`feat:`, `fix:`, `docs:`, `chore:`…). Squash before merge unless repo enforces otherwise.
- **No secrets/PII**: Never include secrets, tokens, or real user data in PR text, code, or artifacts. Use redaction or mocks.
- **Cross‑links**: Link issues/tickets with closing keywords where supported (`Closes #123` / `Fixes #123`).
- **Evidence required**: Include test plan, results, and UI/media evidence where applicable. Summarize perf/security impact.
- **Docs**: Update README, CHANGELOG, MIGRATION, API docs as needed. Breaking changes must include migration notes.
- **Labels & reviewers**: Apply ownership labels, risk/size tags, and request CODEOWNERS early.
- **English**: PR text, templates, and artifacts must be in English unless policy dictates otherwise.

**Platform Notes**

- **GitHub**: Use `.github/pull_request_template.md` and `gh` CLI when available.
- **GitLab**: Use `/.gitlab/merge_request_templates/` and `glab` CLI. Use `Closes #ID` or `Closes GROUP/PROJECT#ID`.
- **Bitbucket**: Templates via default text or repository PR template; link Jira issues with `PROJECT-123`.

## 3. Examples

### 3.1 Good PR Title

```
feat(search): add typo‑tolerant prefix matching and cache warmup

```

### 3.2 Good PR Description (compressed)

```
## What
Adds typo‑tolerant prefix matching to search with a small LRU in front of the DB. Introduces feature flag `search_typo_v1`.

## Why
Users frequently misspell queries; funnel analysis shows 9.8% drop‑off from zero‑results pages.

## How
- New tokenizer and prefix index
- LRU cache (1k entries)
- Flag‑guarded rollout (off by default)

## Tests
- Unit: 142 added/updated (all green)
- E2E: 12 scenarios updated
- Perf: 95th pctl latency +3.1ms (baseline 41.7ms)

## Risks / Rollback
- Risk: cache staleness → mitigated by 60s TTL
- Rollback: flip flag off; revert commit `abc1234` if needed

## Screenshots
[gif] Search before/after (redacted)

## Links
Closes #1234; Design: ADR‑007; Dashboard: Grafana panel 8

## Release notes
Improved search with typo tolerance (flagged)

```

### 3.3 Bad PR Description (anti‑example)

```
fix stuff
- changed files
- tests run

```

### 3.4 Conventional Commit Messages

```
feat(auth): support PKCE for mobile clients
fix(api): return 400 for invalid cursor instead of 500
chore(ci): enable required status checks for PRs

```

### 3.5 CHANGELOG Entry (Keep a Changelog)

```
## [1.12.0] - 2025-09-07
### Added
- Search typo tolerance behind `search_typo_v1` feature flag

### Changed
- API: `/v1/search` now accepts `typo_tolerance` query param

### Deprecated
- None

### Fixed
- Cursor validation on `/v1/list`

```

## 4. Reasoning & Decision Checklist (for Devin)

Use this checklist to make decisions without exposing hidden reasoning:

- Is the PR **single‑scope** and reasonably sized (≤ ~400 net LOC change or split)?
- Does the title follow **imperative mood** and possibly Conventional Commits?
- Does the description answer **What/Why/How**, include **Risks/Rollback**, **Tests**, **Screenshots/GIFs** if UI?
- Are all **linked issues and docs** referenced and reachable?
- Are **breaking changes** flagged and accompanied by **MIGRATION.md** updates?
- Are **labels**, **reviewers**, and **assignees** set per ownership?
- Are CI checks **green** and **coverage** acceptable (threshold per repo)?
- Are secrets/PII absent from code, logs, and media?

## 5. Output Specification

A successful run produces:

- A branch `devin/pr-docs-<slug>` pushed to remote
- Updated documentation:
    - PR description populated from template
    - `CHANGELOG.md` updated (if user‑visible change)
    - `MIGRATION.md` updated (if breaking change)
    - README/API docs updated (if interfaces changed)
- Attached artifacts: redacted screenshots/GIFs, test results summary
- Applied labels (e.g., `type:feature`, `area:search`, `risk:low`, `size:S/M/L`)
- Requested reviewers (CODEOWNERS + domain owners)
- A draft PR converted to “Ready for review” after CI passes

## 6. Procedure

### 6.1 Inputs (set defaults if absent)

- `TARGET_BRANCH` (default `main`)
- `ISSUE_KEY` (e.g., `#1234` or `PROJ-1234`)
- `SLUG` short, kebab‑case description (e.g., `improve-search-typos`)
- `PLATFORM` one of `github|gitlab|bitbucket` (auto‑detect if repo remotes reveal host)

### 6.2 Prepare Branch

1. `git fetch origin --prune`
2. `git checkout -B devin/pr-docs-${SLUG} origin/${TARGET_BRANCH}`
3. Implement the change (or if documenting an existing change, collect commits to include).

### 6.3 Local Validation

1. Run linters/formatters and fix issues.
2. Run unit/integration/e2e tests; collect summaries.
3. For UI changes, capture **redacted** before/after screenshots or a short GIF.
4. Generate or update any API/SDK docs (e.g., `npm run docs`, `sphinx-build`, `docusaurus build`).
5. Update `CHANGELOG.md` (Keep a Changelog format). For breaking changes, create/update `MIGRATION.md` with step‑by‑step instructions and rollback.

### 6.4 Ensure a PR Template Exists

- **GitHub**: Create or update `.github/pull_request_template.md` using the template below.
- **GitLab**: Add `/.gitlab/merge_request_templates/Standard.md` and reference it when creating the MR.
- **Bitbucket**: Configure repository default PR description (admin) or paste from template each time.

### 6.5 Stage and Commit

1. `git add -A`
2. Use Conventional Commits where possible, e.g.:
    - `git commit -m "feat(search): add typo‑tolerant prefix matching (flagged)"`
3. `git push -u origin devin/pr-docs-${SLUG}`

### 6.6 Open PR/MR (choose one)

- **GitHub (gh)**
    
    ```
    gh pr create \  --base ${TARGET_BRANCH} \  --head devin/pr-docs-${SLUG} \  --title "feat(search): add typo‑tolerant prefix matching" \  --body-file .github/pull_request_auto.md \  --draft
    ```
    
    (Create `.github/pull_request_auto.md` from the template and fill placeholders.)
    
- **GitLab (glab)**
    
    ```
    glab mr create \  --source-branch devin/pr-docs-${SLUG} \  --target-branch ${TARGET_BRANCH} \  --title "feat(search): add typo‑tolerant prefix matching" \  --description-file .gitlab/merge_request_auto.md \  --draft
    ```
    
- **Bitbucket**: Open the PR via web UI with the template pasted; set as draft if CI pending.
    

### 6.7 Populate Description & Attach Evidence

1. Fill **What/Why/How**, **Tests**, **Screenshots/GIFs**, **Risks/Rollback**, **Links**, **Release notes**.
2. Add closing keywords: `Closes ${ISSUE_KEY}` (or Jira key mention for Bitbucket/Jira).
3. Upload media or link to artifacts stored in repo under `docs/media/` (preferred) to keep links durable.

### 6.8 Labels, Reviewers, and Rules

1. Apply size/risk labels per repo policy (e.g., `size:S` ≤ 200 LOC, `M` ≤ 600, else `L`).
2. Add domain labels (e.g., `area:search`, `platform:ios`).
3. Request reviewers from `CODEOWNERS` and feature owners.
4. Ensure required status checks and approvals are configured and visible.

### 6.9 Finalize

1. Convert from **Draft** to **Ready for review** once CI is green and description complete.
2. Respond to review comments; update PR text when plans or evidence change.
3. Squash & merge or follow repo’s merge strategy after approvals.

## 7. Advice, Pitfalls, and Quality Gates

**Advice**

- Keep titles short, imperative, and searchable; lead with the change type.
- Prefer short sections with bullets over long paragraphs.
- Use feature flags for risky work; document flag name and default.
- Keep media lightweight (GIF ≤ 10–15s, redacted).

**Common Pitfalls**

- Missing test plan and rollback procedure
- Unlinked issues or dead links to docs/dashboards
- Over‑large PRs without explanation or plan to split
- Screenshots that leak PII or environment details
- CHANGELOG not updated for user‑visible changes

**Quality Gates (must pass)**

- CI green, linters clean, coverage within threshold
- Description complete per template; links valid
- All required labels and reviewers set; approvals met
- No secrets/PII in changes or artifacts
- For breaking changes: `MIGRATION.md` provided and referenced

## 8. Forbidden Actions

- Bundling unrelated changes or refactors without justification
- Bypassing CI or merging red builds
- Merging without required approvals
- Force‑pushing after reviews without clear reason and notice
- Including secrets/PII or vendor‑locked links that others cannot access
- Deleting or altering history to hide context

---

## Templates

### A. PR Description Template (all platforms)

```
# WhatConcise summary of the change in 1–3 sentences.# WhyProblem statement and user impact; link metrics or research.# HowBulleted implementation outline; feature flags, config, migrations.# Tests- Unit: summary- Integration/E2E: summary- Manual/Exploratory: steps and results- Perf: baseline vs. new (p50/p95), memory, etc.# Risks / RollbackKnown risk items and mitigations; exact rollback steps (flag/off, revert SHA, data rollback).# Screenshots / Media (redacted)Before/after images or GIFs with captions. Store under `docs/media/` where possible.# LinksIssues, ADRs, dashboards, design docs: `Closes #1234` / `PROJECT-1234`.# Release NotesOne‑line user‑facing summary.
```

### B. GitHub: `.github/pull_request_template.md`

```
## What## Why## How## Tests- [ ] Unit- [ ] Integration/E2E- [ ] Manual/Exploratory- [ ] Performance## Risks / Rollback- [ ] Rollback plan described- [ ] Feature flag noted (name & default)## Screenshots / Media (redacted)## Links- Issue(s):- Design/ADR:- Dashboards:## Release Notes---**Checklist**- [ ] Scope is single, PR size reasonable- [ ] CHANGELOG/README/API docs updated if needed- [ ] Labels applied, reviewers requested- [ ] No secrets/PII in code or media
```

### C. GitLab: `/.gitlab/merge_request_templates/Standard.md`

```
## What / Why / How## Tests## Risks / Rollback## Media## Links (Closes #ID or Closes GROUP/PROJECT#ID)## Release Notes/label ~"type::feature" ~"risk::low"/assign_reviewer @owner1 @owner2
```

### D. CHANGELOG (Keep a Changelog)

```
## [Unreleased]### Added-### Changed-### Deprecated-### Removed-### Fixed-### Security-
```

### E. MIGRATION.md (for breaking changes)

```
# Migration for <Feature/Module>## SummaryWhat changed and why it’s breaking.## Affected APIs/Configs-## Step‑by‑step1.2.## RollbackExact steps to revert (flag, revert commit, data corrections).
```

### F. Dangerfile (optional automation)

```
# fail if PR description is too shortfail("Please complete the PR description.") if github.pr_body.length < 200# warn on large PRswarn("PR is large; consider splitting.") if git.lines_of_code > 600# require changelog if src changed and not docs‑onlyhas_src_changes = !git.modified_files.grep(/^(src|app|lib)\//).empty?changelog_changed = git.modified_files.include?("CHANGELOG.md")fail("Update CHANGELOG.md") if has_src_changes && !changelog_changed
```

---

## Output Format

- A single PR/MR on the chosen platform with:
    - Title using imperative mood (prefer Conventional Commits)
    - Description completed using template
    - Linked issues/epics and docs
    - Attached media artifacts (redacted)
    - Labels and reviewers set
    - Updated docs: CHANGELOG, MIGRATION, README/API where applicable
- All changes submitted in a branch named `devin/pr-docs-<slug>` and merged via standard process.
# Documentation / README Generation — Devin Playbook

## Overview

This playbook defines a repeatable workflow to create or overhaul a project’s **README.md** and lightweight docs so that new users and contributors can install, run, and extend the software with minimal friction. It covers repo analysis, content extraction from code/CI, generation of a structured README, validation (lint + link checks + runnable setup), and delivery via a Pull Request.

---

## Background Data (Intake)

Gather before starting. If missing, infer from repo and add TODOs.

- **Repository**: URL, default branch, license file path
- **Project type**: app/service/lib/CLI/ML project/infra; language(s); framework(s)
- **Runtime matrix**: required OS/es, CPU/GPU, language/runtime versions
- **Build/Run**: package manager(s), entrypoints, Dockerfile(s), compose files
- **Data/Services**: databases, queues, external APIs, cloud services
- **Config**: env vars, secrets (document **names only**), config files
- **Testing**: frameworks, how to run tests, coverage
- **CI/CD**: provider, key workflows, badges
- **Deployment**: targets (k8s, serverless, VM), example configs
- **Docs**: existing README/docs site, style guide
- **Maintenance**: maintainers, contact, support policy
- **Compliance**: privacy, PII handling, security posture

> Intake YAML (drop in /docs/intake.yaml)

```
repo: <https://example.com/org/repo>project_type: servicelanguages: [python]frameworks: [fastapi]runtimes: { python: "3.11" }entrypoints: ["uvicorn app.main:app --reload"]package_manager: pipbuild: ["pip install -r requirements.txt"]env:  required: [DATABASE_URL, SECRET_KEY]  optional: [LOG_LEVEL]services: [postgres]ci: [github-actions]deployment: [docker, kubernetes]docs_style: commonmarkmaintainers:  - name: Jane Doe    email: dev-team@example.comsupport_policy: "Best-effort, business hours PT"
```

---

## Task Description & Rules

**Goal:** Produce a complete, accurate `README.md` (and small supporting docs) that:

- Is runnable end-to-end by a new developer following the instructions.
- Documents features, architecture, configuration, common tasks, and contribution workflow.
- Includes badges, TOC, consistent formatting, and valid links.

**Rules**

1. **No secrets** in docs or examples. Reference names of env vars only.
2. **Do not fabricate** features or APIs. If uncertain, mark `TODO:` and open a question in the PR.
3. Prefer **copying truth from code/CI** (package.json, pyproject.toml, Dockerfile, workflows) over guessing.
4. Keep sections **scannable**; use headings, bullets, and code blocks.
5. Use **relative links** for intra-repo references.
6. Ensure **setup commands actually run** in the VM (or are marked as TODO with reason).
7. Keep README **project-focused**; move long guides to `/docs/`.
8. End work with a **PR** from a `devin/` branch with evidence of validation.

---

## Output Format (Deliverables)

Create/modify the following files:

- `README.md` — canonical entrypoint (required)
- `CONTRIBUTING.md` — contribution guide (recommended)
- `docs/` — optional pages: `architecture.md`, `operations.md`, `FAQ.md`
- `.markdownlint.json` — lint config (optional)
- `.github/workflows/docs-ci.yml` — markdown lint + link check (optional)

**Naming/Style**

- Title case headings, fenced code blocks, shell commands prefixed with `$`.
- Keep lines ≤ 100 chars where reasonable.
- Badges at top (build, coverage, license, version).

---

## Step-by-step Thinking (Reasoning Checklist)

Use this lightweight checklist to guide decisions:

1. **Identify audience** (user vs contributor) and prioritize their top tasks.
2. **Discover truth** from repo artifacts (package manifests, CI, Docker, compose, scripts).
3. **Minimize cognitive load**: defaults first, advanced later.
4. **Make it runnable**: every command block should be executable as-is or clearly marked.
5. **Prove it works**: run setup/test locally; capture actual outputs or versions.
6. **Guard rails**: add warnings for destructive commands; avoid `sudo` unless required.
7. **Automate quality**: lint markdown, check links in CI.

---

## Procedure

### 1) Repo Intake & Scan

1. Create branch: `devin/docs-readme`.
2. Clone repo; list top-level files/folders.
3. Parse language/package files for metadata:
    - Node: `package.json` (name, scripts, engines)
    - Python: `pyproject.toml` / `requirements.txt`
    - Java: `pom.xml` / Gradle
    - Rust: `Cargo.toml`
4. Inspect `Dockerfile*`, `docker-compose*.yml`, `Procfile`, `Makefile`, `/scripts/`.
5. Inspect CI (e.g., `.github/workflows/*.yml`) for build/test/runtime versions.
6. Inventory env vars referenced in code and compose/manifests.
7. Open `LICENSE`, `CODE_OF_CONDUCT.md`, existing docs.

### 2) Draft Structure

Create a minimal but complete scaffold:

```
# <Project Name>[![CI](<badge>)](<link>) [![Coverage](<badge>)](<link>) [![License: MIT]](LICENSE)## OverviewOne-paragraph purpose and key features.## Quickstart- Prereqs- Install- Run## ConfigurationEnv vars and config files.## UsageCommon commands / API examples.## ArchitectureHigh-level diagram + key components.## DevelopmentLocal dev commands, tests, lint, formatting.## DeploymentDocker/k8s or platform-specific notes.## TroubleshootingCommon issues and fixes.## ContributingLink to CONTRIBUTING.md.## LicenseSPDX + link.
```

### 3) Extract Truth & Fill Content

1. Fill **Quickstart** using actual commands from scripts/Makefile/compose.
2. Document **Configuration** by listing env var names, defaults, and impact.
3. Add **Usage** with realistic examples (CLI usage, HTTP endpoints, or code snippets) sourced from tests/examples.
4. Summarize **Architecture**; include a Mermaid diagram (see template below).
5. Record **Development** tasks: `install`, `format`, `lint`, `test`, `coverage`.
6. Add **Deployment** based on Docker/k8s manifests; note image name, ports, health checks.
7. Add **Badges** (CI, coverage, license). If unavailable, omit or add TODOs.

### 4) Validate by Running Locally (VM)

1. Follow your Quickstart exactly. Capture any missing steps.
2. Run tests: ensure they pass or document known failures.
3. Run a **link check** over README and docs.
4. Run **markdown lint**. Fix headings, code fences, line length as needed.

**Commands (examples)**

```
# link check (Node)$ npx markdown-link-check README.md -q# markdown lint$ npx markdownlint-cli2 "**/*.md"# or Python alternative$ pip install mdformat$ mdformat --check README.md || mdformat README.md
```

### 5) Wire Up Docs CI (Optional but Recommended)

Add `.github/workflows/docs-ci.yml` (template below) to lint and check links on PRs.

### 6) Finalize & Open PR

1. Update screenshots or sample outputs if applicable.
2. Ensure `README.md` is the single source of truth; deep content goes to `/docs/`.
3. Commit with conventional messages.
4. Open PR with evidence of validation.

---

## Validation & Acceptance Criteria

- [ ]  `README.md` contains: Overview, Quickstart, Configuration, Usage, Architecture, Development, Deployment, Troubleshooting, Contributing, License.
- [ ]  All commands are executable or marked with TODO and rationale.
- [ ]  `markdownlint` passes locally.
- [ ]  `markdown-link-check` passes or false-positives are ignored via config.
- [ ]  Local run of the service/lib succeeds (or limitations documented).
- [ ]  No secrets or tokens committed; env names only.
- [ ]  Badge links resolve or are omitted intentionally.
- [ ]  PR includes screenshots/logs, test results summary, and checklists above.

---

## PR & Branching

- **Branch name**: `devin/docs-readme` (or `devin/readme-overhaul` for large changes)
- **Commit examples**:
    - `docs(readme): add quickstart and config table`
    - `docs: add architecture diagram and dev workflow`
    - `ci(docs): add markdown lint + link check`
- **PR description template**:
    - **Summary**: What changed and why
    - **Scope**: Files touched
    - **How I validated**: commands run, outputs, screenshots
    - **Open questions/TODOs**: list blockers
    - **Checklist**: copy of Acceptance Criteria

---

## Advice (Good Practices)

- Prefer **copy/paste** from scripts over paraphrasing to prevent drift.
- Put **the shortest path to “it runs”** near the top.
- Provide **copyable blocks**; avoid inline prompts like “replace X with Y” inside code.
- Use **tables** for env vars and ports.
- Add **support policy** (what versions/platforms you test against).
- Keep README focused; create `/docs/architecture.md` for deeper dives.
- When in doubt, **show real examples** (CLI usage, API curl, code snippets).

---

## Forbidden Actions

- Revealing secrets, tokens, or credentials.
- Inventing features or endpoints not present in the code.
- Deleting unrelated files or refactoring code as part of documentation work.
- Running destructive commands against production resources.
- Using `sudo` or system-level changes unless absolutely necessary and documented.
- Changing license or legal text without explicit direction.

---

## Templates & Snippets

### README.md Template

```
# <Project Name>[![Build](<badge>)](<link>) [![Coverage](<badge>)](<link>) [![License]](LICENSE)> One-paragraph elevator pitch. Who is this for? What does it do?## Table of Contents- [Quickstart](#quickstart)- [Configuration](#configuration)- [Usage](#usage)- [Architecture](#architecture)- [Development](#development)- [Deployment](#deployment)- [Troubleshooting](#troubleshooting)- [Contributing](#contributing)- [License](#license)## Quickstart**Prerequisites**- OS: macOS/Linux/Windows- Runtime: <e.g., Node 20 / Python 3.11>- Tools: <e.g., Docker, Make>**Setup**```bash$ git clone <repo>$ cd <repo>$ <install command>$ <run command>
```

**Verify**

```
$ <health or test command>
```

## Configuration

|Name|Required|Default|Description|
|---|---|---|---|
|EXAMPLE|yes|—|What it controls|

## Usage

```
$ <example command>
```

HTTP example:

```
$ curl -i http://localhost:8080/health
```

## Architecture

```
graph TD  client[Client] --> api[API]  api --> svc[Service]  svc --> db[(DB)]
```

Key components and data flows.

## Development

```
$ <install deps>$ <format>$ <lint>$ <test>$ <coverage>
```

## Deployment

- Docker image: `<org/image:tag>`
- Ports: `8080/tcp`
- Health: `/healthz`

## Troubleshooting

- Symptom → Cause → Fix

## Contributing

See [CONTRIBUTING.md](https://chatgpt.com/g/g-68af6c97fa608191904c6f89616901d0-playbookgpt/c/CONTRIBUTING.md).

## License

SPDX: MIT — see [LICENSE](https://chatgpt.com/g/g-68af6c97fa608191904c6f89616901d0-playbookgpt/c/LICENSE).

````

### CONTRIBUTING.md (minimal)
```md
# Contributing

## Development Setup
- Follow README Quickstart.
- Create branches: `devin/<short-purpose>`.

## Coding Standards
- Format + lint before pushing.
- Write tests for new features.

## Pull Requests
- Small, focused changes.
- Include screenshots/logs for docs updates.
- Link related issues.

````

### Docs CI Workflow (`.github/workflows/docs-ci.yml`)

```
name: Docs CIon:  pull_request:    paths: ["**.md", ".github/workflows/docs-ci.yml"]jobs:  lint-and-links:    runs-on: ubuntu-latest    steps:      - uses: actions/checkout@v4      - uses: actions/setup-node@v4        with: { node-version: 20 }      - run: npm -g i markdownlint-cli2 markdown-link-check      - run: markdownlint-cli2 "**/*.md"      - run: |          set -e          find . -name "*.md" -print0 | xargs -0 -n1 markdown-link-check -q || true
```

### Markdownlint Config (`.markdownlint.json`)

```
{  "default": true,  "MD013": { "line_length": 100, "code_blocks": false },  "MD033": false,  "MD041": false}
```

### Link Check Ignore (optional `.mlc_config.json`)

```
{  "ignorePatterns": [    { "pattern": "^http://localhost" },    { "pattern": "^https://example.com" }  ]}
```

---

## Troubleshooting

- **Install fails** → check runtime versions from CI config; pin versions in README.
- **Docker build errors** → copy exact build command from CI or Makefile; ensure context path.
- **Link checker noisy** → ignore localhost and non-deterministic links via config.
- **Monorepo** → top-level README links to package-level READMEs; add a workspace map.

---

## Special Cases & Notes

- **Libraries**: emphasize install, minimal example, API surface, versioning/semantic release.
- **Services**: emphasize run, health, configuration, ports, deployment.
- **CLI tools**: show `-help`, subcommands, exit codes, shell completion.
- **ML projects**: include data sources, experiment tracking, model artifacts, GPU notes.
- **Infra/IaC**: include plan/apply workflows, state mgmt, providers, policy.

---

## Completion Criteria (What “Done” Means)

- README is **accurate**, **runnable**, **linted**, and **link-clean**.
- Optional docs added for architecture/operations as needed.
- CI in place for ongoing docs quality.
- Work delivered via PR with validation evidence.
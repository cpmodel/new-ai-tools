# AGENTS.md

Instructions for AI agents working in this repository. `CLAUDE.md` points here.

## Overview

Fiscal Climate AI Tools: working papers, data workbooks, and process files for
fiscal and climate-policy analysis, organized by project.

- `README.md` — repository structure and project summaries.
- `SETUP.md` — environment requirements, build steps, and run-process specs.
- `CPATLinks/`, `ExciseDiagnostic/`, `TaxMacroFiscalDashboard/` — project folders.

Read `README.md` and `SETUP.md` before making changes, and any process
specification inside the project folder you are touching.

## Working rules

- Keep changes scoped to the task; don't reorganize or rename files unprompted.
- Follow the project's process specifications (e.g. `process_files/*.md`) when
  running or updating an analysis; don't deviate silently.
- Treat data workbooks (`.xlsx`) and working papers (`.docx`) as source
  artifacts: don't overwrite them without being asked, and note which version
  you edited.
- Never invent data, figures, or citations. Cite sources and flag assumptions
  or gaps clearly.
- Keep documentation (`README.md`, `SETUP.md`) in sync when structure,
  requirements, or run steps change.
- Don't commit secrets, credentials, or large generated outputs.

## Git

- Develop on the branch you were assigned; don't push elsewhere.
- Write clear, descriptive commit messages.
- Open pull requests only when asked.

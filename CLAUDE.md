# CLAUDE.md

This file provides guidance for AI assistants (Claude and others) working in this repository.

## Repository Overview

**Name:** Default
**Status:** Newly initialized — no source code, dependencies, or tooling have been added yet.
**Remote:** `http://local_proxy@127.0.0.1:42533/git/cgs5hjjh7y-ops/Default`

The repository currently contains only this file and a placeholder `README.md`. All project structure, conventions, and tooling should be established as development begins.

## Current State

```
/
├── .git/
├── CLAUDE.md          ← this file
└── README.md          ← placeholder ("# Default")
```

## Development Branch

Active development happens on branches prefixed with `claude/`. Always push to the designated feature branch, never directly to `master` without explicit instruction.

## Starting a New Project Here

When a language, framework, or project type is chosen, update this file with:

1. **Project type and purpose** — what it does, who uses it
2. **Tech stack** — language(s), frameworks, major libraries
3. **Directory layout** — where source, tests, configs live
4. **How to install dependencies** — e.g. `npm install`, `pip install -r requirements.txt`
5. **How to run the project** — dev server, entry point
6. **How to run tests** — test command, coverage command
7. **How to build** — build/compile command if applicable
8. **Linting and formatting** — tools used and how to run them
9. **Environment variables** — what's required, `.env.example` location
10. **Key conventions** — naming, code style, commit message format

## General Conventions (to be refined when project is defined)

### Git

- Commit messages should be short, imperative, and descriptive (e.g. `Add user auth`, `Fix null pointer in parser`)
- Prefer small, focused commits over large monolithic ones
- Feature branches should be merged via pull request

### Code Quality

- Prefer simple, readable code over clever abstractions
- Only add comments where the logic is not self-evident
- Do not add dead code, unused imports, or unnecessary type casts
- Handle errors explicitly; do not silently swallow exceptions

### File Hygiene

- Do not commit `.env` files, secrets, or credentials
- Do not commit build artifacts, compiled outputs, or `node_modules`/`venv`-style dependency directories
- Keep the root directory clean — configuration files belong at the root, source code in a subdirectory

## Notes for AI Assistants

- This repository is empty. Do not invent or assume a project structure.
- When asked to implement something, ask about the tech stack first if it has not been established.
- Update this `CLAUDE.md` file whenever significant project decisions are made (tech stack chosen, major libraries added, conventions established).
- Before making changes, read existing files rather than assuming their contents.
- Do not create files speculatively — only create what is needed for the current task.

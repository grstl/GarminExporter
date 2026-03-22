# AGENTS.md

## Cursor Cloud specific instructions

This repository ("GarminExporter") is currently a greenfield project with no source code, dependencies, or build system. The only file is `README.md`.

### Current state

- **No language/framework chosen yet** — no `package.json`, `requirements.txt`, `pyproject.toml`, `go.mod`, or similar.
- **No services to run** — no application entry point, no dev server, no database.
- **No tests or linters configured.**
- **No Docker/devcontainer setup.**

### When code is added

Once source code and a dependency manifest are introduced, update this section and the VM update script accordingly. Key things to configure:
- Dependency installation command (e.g. `pip install -r requirements.txt`, `npm install`)
- Dev server startup instructions
- Test and lint commands
- Any required secrets or environment variables (e.g. Garmin Connect credentials)

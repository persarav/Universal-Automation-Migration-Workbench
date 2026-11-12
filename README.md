# Universal Automation Migration Workbench

A local-first web application prototype for understanding source code/RPA workflows and preparing migration packages for another language or RPA platform.

## What is included

- Upload source files or a ZIP project: `.xlsm`, `.xlam`, `.bas`, `.cls`, `.frm`, `.py`, `.vb`, `.vbs`, `.cs`, `.xaml`, `.json`, `.xml`, `.flow`, `.atmx`, `.bprelease`, `.md`.
- Inspect source files, symbol/function inventory, source line numbers, call hints, decision/loop/error indicators and external-system touchpoints.
- Read macro-enabled Excel workbook metadata and extract VBA modules with `oletools`. **No macros are executed.**
- Parse basic UiPath XAML activities and JSON-based flow definitions.
- Build a platform-neutral intermediate representation (IR) before target generation.
- Import Agent Skills-style skill packs from a GitHub repository or ZIP. Indexes `SKILL.md` and Markdown reference docs only; imported scripts are not executed.
- Generate a migration report, traceability table, unresolved-items list, regression test plan and target scaffold.
- Optional AI-assisted semantic conversion through an OpenAI-compatible Chat Completions API. Source transfer is disabled unless the user explicitly checks the consent box.

## Run on Windows

1. Install Python 3.10 or newer.
2. Extract this project folder.
3. Double-click `run_windows.bat`, or run these commands in PowerShell:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

4. Open the local Streamlit URL shown in the terminal (normally `http://localhost:8501`).

## Run with Docker

```bash
docker build -t migration-workbench .
docker run --rm -p 8501:8501 migration-workbench
```

## Optional AI configuration

- Enter the API key, model name and OpenAI-compatible base URL in the Migrate page; or set environment variables `OPENAI_API_KEY`, `MIGRATION_MODEL`, and `OPENAI_BASE_URL`.
- The default model name is `gpt-4.1`; change it to a model available to your endpoint.
- The model call happens only when the user selects AI-assisted generation **and** checks the source-sharing consent box.
- Common credential assignment and bearer-token patterns are best-effort redacted before source text is sent to the model. This is not a complete secret scanner; inspect the source and rotate any exposed secret.
- The app does not save the API key in the downloadable migration package.
- Review your employer's code/data handling policy before sending source code, workbook content or skill references to an external model.

## Skills

The Skill Library page catalogues the supplied repositories and allows importing skill packages from GitHub or ZIP. It fetches `SKILL.md` and Markdown references only. It deliberately does not execute third-party code or fetch arbitrary executable assets into the prompt.

The Python topic URL is a GitHub discovery page, not a single skill repository. Choose specific repositories from that topic before importing. Check the upstream license, version and maturity before adopting a skill pack.

## Important limitations

This is a migration workbench prototype, not a universal compiler. Direct code-to-code or RPA-to-RPA conversion cannot be guaranteed for all products. A valid target migration often requires tenant metadata, connector operation schemas, activity package versions, selectors, credentials, queues, assets, machine configuration, proprietary export formats and business-owner decisions.

- The deterministic mode provides static facts and clearly labeled scaffolds. It does not infer complete business semantics.
- The LLM mode can infer and generate more, but generated output still requires human review and native target validation.
- The workbench is designed for local use or deployment behind an organization-approved authentication gateway. Do not expose an unauthenticated Streamlit instance publicly.
- UiPath output from fallback mode is a starter XAML with TODO placeholders, not a compiled workflow.
- Power Automate, Automation Anywhere and Blue Prism fallback output is a design blueprint, not an import-ready package.
- Excel VBA extraction uses `oletools`; the app never runs VBA macros.
- The current workbook parser extracts macro code and basic sheet/header metadata, but does not fully emulate workbook buttons, forms, ActiveX, formulas, events or external connections.
- The `.atmx` and `.bprelease` extension support is currently file ingestion and evidence inspection, not complete proprietary package decompilation.

## Quality gates for real migrations

1. Inventory source files, modules, triggers and dependencies.
2. Reconstruct business rules and all input/output contracts with source evidence.
3. Review the IR with the process owner.
4. Map actions to target-native activities/connectors and verify schemas.
5. Bind credentials, selectors, assets and environment resources securely.
6. Generate target files and run the native target validator/build.
7. Execute regression tests comparing outputs and side effects with the original.
8. Complete security, operations, rollback and deployment approval.

## Project layout

```text
app.py             Streamlit web UI
engine.py          Source extraction, static analysis, IR and target scaffolds
skill_loader.py    Safe Agent Skills/GitHub repository importer
requirements.txt   Python dependencies
Dockerfile         Container deployment
run_windows.bat    Windows local startup script
tests/             Unit tests
examples/          Small source samples
```

## Tests

```bash
pytest -q
```

## Skills and repositories referenced

- Microsoft Power Automate skills: https://github.com/microsoft/power-platform-skills/tree/main/plugins/power-automate
- UiPath skills: https://github.com/UiPath/skills
- VBA skills: https://github.com/RenRMT/claude-code-vba-skills
- Python skills topic: https://github.com/topics/python-skills
- Agent Skills: https://github.com/agentskills/agentskills and https://agentskills.io/specification

The upstream repositories are maintained by their respective owners and can change independently of this project.

## Free cloud deployment

See [DEPLOYMENT.md](DEPLOYMENT.md) for the Streamlit Community Cloud deployment steps. For source-code privacy, use a private GitHub repository and keep the Streamlit app private. Do not commit API keys or customer source files.

# Universal Automation Migration Workbench — Project Files

This package contains the Streamlit application source, deterministic analysis engine, UiPath XAML generator, skill loader, VBA extractor, bundled skill Markdown packs, tests, deployment configuration and documentation.

## Main files
- `app.py` — Streamlit interface and workflow orchestration.
- `engine.py` — source parsing, normalized workflow model and deterministic analysis.
- `uipath_generator.py` — UiPath project/XAML generation.
- `vba_extractor.py` — VBA extraction from macro-enabled Excel files.
- `skill_loader.py` — local/GitHub skill-pack loading.
- `skillpacks/` — bundled platform skill Markdown files.
- `tests/` — automated tests.
- `requirements.txt` — Python package dependencies.

## Run locally
1. Install Python 3.10 or later.
2. Create and activate a virtual environment.
3. Install dependencies: `python -m pip install -r requirements.txt`.
4. Start the app: `streamlit run app.py`.

An AI API key is optional for the deterministic/offline analysis path. AI-assisted interpretation requires a configured model endpoint.

## Streamlit Community Cloud
Push the contents of this package to the root of a GitHub repository. In Streamlit Community Cloud, select that repository, branch `main`, and main file path `app.py`.

- `tools/validate_uipath_project.py` — static preflight validator for generated UiPath Windows XAML projects (does not replace Studio Validate/Build).

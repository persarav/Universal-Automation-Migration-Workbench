# Deploy Universal Automation Migration Workbench for free

## Recommended host: Streamlit Community Cloud

This project is already a Python/Streamlit application, so Streamlit Community Cloud is the shortest free deployment route. It provides a `*.streamlit.app` URL and deploys from a GitHub repository.

## Privacy recommendation

Use a **private GitHub repository** and keep the Streamlit app **private**. The workbench accepts VBA, Excel workbooks, source code, and RPA project files. Do not make it public until you have reviewed the source-handling and access requirements. Streamlit Community Cloud documents that an app deployed from a private repository is private by default; viewers can be invited explicitly. The free Community Cloud account allows only one private app at a time, according to its current documentation.

## Files to upload to GitHub

Upload the contents of this folder to the repository root. Do not upload:

- `.env` files or real API keys
- `.streamlit/secrets.toml`
- `__pycache__/`, `.pytest_cache/`, `.venv/`
- uploaded business source files, exported reports, or test workbooks containing real data

The repository includes a `.gitignore` to exclude local secrets and caches.

## Deploy steps

1. Sign in to GitHub and create a **private** repository, for example `universal-automation-migration-workbench`.
2. Upload the project files from this folder to the repository root, or push them with Git.
3. Open <https://share.streamlit.io/> and sign in using the same GitHub account.
4. Select **Deploy an app**, then choose the private repository, branch `main`, and entrypoint `app.py`.
5. Choose an available app subdomain. Suggested name: `universal-migration-workbench`. The resulting URL is expected to be `https://universal-migration-workbench.streamlit.app` if that name is available; the provider confirms the actual URL during deployment.
6. Keep app privacy set to **Only specific people can view this app**. Invite additional users by email from the app's Share settings if needed.
7. Wait for the provider's build to complete, then test upload, VBA extraction, static analysis, fallback generation, and download.

Official instructions:
- <https://docs.streamlit.io/deploy/streamlit-community-cloud>
- <https://docs.streamlit.io/deploy/streamlit-community-cloud/get-started>
- <https://docs.streamlit.io/deploy/streamlit-community-cloud/share-your-app>

## Dependencies and entrypoint

- Entrypoint: `app.py`
- Dependency file: `requirements.txt`
- Python: select Python 3.11 in the app's advanced settings if the platform asks for a version.

## AI provider configuration

AI-assisted generation requires a provider API key. Do not commit a key to GitHub. The safest initial deployment is to leave AI generation disabled until you add the key through the hosting platform's secrets settings. The UI requires the user to explicitly approve source transfer before calling the configured model.

Optional Streamlit secrets file (create in the hosting dashboard, not in GitHub):

```toml
OPENAI_API_KEY = "<your-provider-key>"
MIGRATION_MODEL = "<model-supported-by-your-provider>"
OPENAI_BASE_URL = "https://api.openai.com/v1"
```

Only configure this if your organization's code/data policy allows the selected provider to receive source material.

## Post-deployment smoke test

- Confirm the app URL opens and privacy is restricted as intended.
- Upload a harmless `.bas` sample and confirm procedure inventory appears.
- Upload a sample `.xlsm` with no sensitive data and confirm macro extraction; uploaded macros are not executed.
- Generate the deterministic scaffold without AI.
- Download and inspect the migration ZIP.
- Test AI mode only after configuring a provider key and confirming the source-sharing consent control works as expected.

## Known limitations

This is a migration workbench prototype, not a universal compiler. Native RPA export formats and production-ready migrations still require platform-specific parsers, activity/connector metadata, credentials and environment bindings, native validation, and regression testing. Do not treat generated scaffolds as production-ready bots.


## Updating an existing Streamlit deployment

1. Download and extract the latest deployment ZIP.
2. In the connected GitHub repository, replace the existing root files with the updated root files from the ZIP (do not upload the ZIP itself).
3. Ensure `app.py`, `engine.py`, `uipath_generator.py`, `skill_loader.py`, `requirements.txt`, `.streamlit/config.toml` and the complete `skillpacks/` folder are committed at the repository root.
4. Commit to the branch configured in Streamlit (normally `main`). Streamlit Community Cloud should detect the commit and rebuild the app.
5. Open Manage app → Logs if the rebuild fails. Share the redacted traceback, not secrets or source workbooks.

No API key is required for source parsing, the bundled skill library, the SYNC-specific offline UiPath generation, or the static project validator. API/model configuration is optional.

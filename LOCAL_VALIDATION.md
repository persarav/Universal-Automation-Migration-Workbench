# Local validation record

Validation performed in the build environment on 2026-11-12.

## Passed

- `python -m py_compile app.py engine.py uipath_generator.py skill_loader.py build_sync_project.py validate_uipath_project.py` — passed.
- `pytest -q` — **16 tests passed**.
- Actual `SYNC_password_removed.xlsm` source extraction — 19 VBA source files, including `ThisWorkbook` event code.
- Source inventory — **249 Sub/Function procedures detected** and **249 mapping rows generated**.
- UiPath generator — **20 XAML files** generated: one Main orchestrator, 17 cohesive business/utility workflows and 2 SAP adapter workflows.
- Bundled offline skills — **9 Markdown playbooks** included in the app and standalone UiPath project.
- XAML parsing — every generated XAML file parsed as well-formed XML.
- Static UiPath project validator — passed: project manifest, workflow references, unique activity IDs, argument bindings, mapping targets and package references were checked.
- No placeholder `TODO: implement` activities are emitted by the SYNC-specific generator.
- Offline generation — `uses_model_api_key=false`.
- Procedure mapping regression checks — snapshots, retired PO helpers, STF candidate logic, appointment parsing, SAP inventory extraction and dashboard-only actions map to the intended owning workflows.
- Migration result normalization — `None`, non-dictionary values, malformed file content and malformed mapping rows are normalized before Streamlit rendering.

## Important limits (not verified in this environment)

- Streamlit is not installed in the build container, and outbound package-index access is unavailable, so the web server could not be launched here. The deployment runtime must install `requirements.txt`.
- UiPath Studio is not installed here. Static XML validation is **not** equivalent to UiPath Studio Validate/Build; the project must be opened and built in the intended Studio version.
- SAP GUI selectors for ME2M, ZPRR_INVSTAT and ME23N cannot be inferred safely from VBA. The generated SAP adapters intentionally fail closed until selectors are captured in the real SAP environment.
- Outlook draft/send requires a configured, approved UiPath Outlook or Microsoft 365 activity and mailbox connection. The generated Email_Analysts workflow builds the recipient/CC/subject/body/attachment-row data but does not send mail.
- The nine bundled Markdown files are curated offline playbooks. The Skill Library can fetch `.md`/`.mdx` files from the supplied GitHub repositories without an AI API key (subject to GitHub rate limits), then include them in the export. Python and Agentic Skills GitHub topic pages are discovery indexes, not canonical repositories.
- For this SYNC workbook, deterministic business logic is generated offline. For arbitrary unrecognized source projects, the generic path remains an analysis/design scaffold until a source/target adapter is built or optional model assistance is enabled.

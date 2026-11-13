# Local validation record

Validation performed in the build environment on 2026-11-12.

## Passed

- Python compilation: `python -m py_compile app.py engine.py uipath_generator.py skill_loader.py vba_extractor.py` — passed.
- Automated tests: `pytest -q` — **17 tests passed**.
- Actual `SYNC_password_removed.xlsm` extraction: **19 VBA source modules** extracted without running macros.
- Source inventory: **249 Sub/Function procedures** detected.
- Actual offline UiPath generation: **20 XAML workflows/adapters** generated and **249 source-to-target mapping rows** created.
- XAML XML parsing: all 20 generated XAML files parsed successfully.
- UiPath Studio 24.10.8 generator regression: generated XAML contains no `x:DateTime`, no `ue:` prefix, includes the correct `sys:DateTime` CLR namespace, and uses `ui:ReadRange` / `ui:WriteRange`.
- Static UiPath project validator: passed on the project generated from the actual SYNC workbook. Checks project manifest, target Studio version, XAML root/class names, type aliases, argument declarations, unique activity IDs, invoked workflow paths/contracts, Invoke Code argument bindings, Workbook activity namespaces, package references and 249-row source traceability.
- Result-rendering normalization: invalid or empty generator results and malformed generated-file/mapping payloads are normalized before Streamlit rendering.
- Offline mode: SYNC-specific migration uses deterministic parsing and XAML generation without an AI/model API key.
- Bundled skills: 7 curated platform/agent skill packs plus one XAML reference document and the skill index are included.

## Important limits (not verified in this environment)

- Streamlit Community Cloud was not launched or rebuilt from this updated package in this environment. Replace the files in the connected GitHub repository and let Streamlit rebuild before treating the hosted app as updated.
- UiPath Studio is not installed here. Static XAML/project validation is **not** equivalent to UiPath Studio Validate/Build; open the generated project in Studio 24.10.8, restore packages, Validate and Build it.
- SAP GUI selectors for ME2M, ZPRR_INVSTAT and ME23N cannot be inferred safely from VBA. The generated SAP adapters intentionally fail closed until selectors are captured in the target SAP environment.
- Outlook draft/send requires configured UiPath Outlook or Microsoft 365 activities and an approved mailbox connection. The generated Email_Analysts workflow builds the analyst email queue and content; it does not send mail.
- The bundled Markdown skill packs are curated offline playbooks, not a full mirror of every upstream repository. The Skill Library can fetch Markdown files from supplied public GitHub repository/subtree URLs without a model API key, subject to GitHub availability/rate limits. Imported scripts/binaries are not executed.
- For arbitrary, previously unseen source/target pairs, the generic path produces a traceable analysis and migration scaffold; it does not claim full executable semantic translation. The SYNC workbook has a dedicated deterministic migration path.

## Re-run local checks

```bash
python -m py_compile app.py engine.py uipath_generator.py skill_loader.py vba_extractor.py
pytest -q
```

To validate a generated UiPath project directory:

```bash
python tools/validate_uipath_project.py PATH_TO_GENERATED_UIPATH_PROJECT
```

This static validator is a preflight only. The authoritative platform gate is UiPath Studio/CLI Validate and Build in the actual target environment.

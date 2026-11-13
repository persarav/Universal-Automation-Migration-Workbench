# Bundled migration skill library

These are locally authored migration playbooks designed for deterministic, no-API-key analysis and generation. They are not verbatim copies of vendor repositories. The app can refresh/import every Markdown file from the upstream public repositories listed in `SKILL_SOURCES.md`; imported Markdown is treated as reference data and scripts are never executed.

Each platform pack identifies its source/target parsing conventions, required runtime details, generation expectations, validation gates and known non-portable capabilities.

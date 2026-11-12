# Skill source review and integration notes

Reviewed against publicly visible repository documentation on 2026-11-12. These are integration references, not bundled copies of the upstream projects.

## Microsoft Power Platform / Power Automate

Source: https://github.com/microsoft/power-platform-skills/tree/main/plugins/power-automate

The Power Automate plugin documents setup, browsing, guided and autonomous flow building, debugging/diagnosis, flow lifecycle management, desktop-flow management and environment routing. It is backed by a FlowAgent MCP engine. Its build guidance emphasizes checking templates first, discovering the connector operation, obtaining exact operation details (including parameter names, enums and action type), validating/preflighting the flow, creating in a stopped state, verifying the created flow and only then publishing. The plugin also covers connection discovery/fixing, dynamic values, backups and run diagnosis. This workbench therefore treats connector metadata and environment IDs as required target evidence; it does not guess action names or parameters.

## UiPath

Source: https://github.com/UiPath/skills

The UiPath skills repository describes authoring, platform operations, diagnostics and lifecycle skills. The `uipath-rpa` skill routes `.xaml` and `.cs` workflows, distinguishes modern Windows/Portable and Legacy target frameworks, prefers XAML for standard RPA unless coded workflows are requested/needed, and routes testing, platform operations, solutions and agents to dedicated skills. The repository README also marks some skills as Stable, Preview or In-development. A migration engine must therefore record project type, target framework, activity/package compatibility and skill maturity rather than generate one generic XAML for every project.

## VBA

Source: https://github.com/RenRMT/claude-code-vba-skills

The visible VBA skill reference emphasizes clear verb-based procedure naming, descriptive variables, explicit data types, enums instead of magic numeric constants, user-defined types and dictionaries/collections for structured data. These conventions are incorporated as review and modernization guidance; they are not a substitute for understanding Excel object-model behavior, workbook state, formulas, macro events or SAP integrations.

## Python skills topic

Source: https://github.com/topics/python-skills

This URL is a topic/discovery page containing multiple independent repositories and is not one coherent canonical skill pack. The application does not silently choose one. Users import a particular repository after reviewing its purpose and license.

## Agent Skills format

Source: https://agentskills.io/specification

The format defines skill folders with a required `SKILL.md` and optional `scripts/`, `references/` and `assets/`. `SKILL.md` contains YAML frontmatter and Markdown instructions; the format defines required `name` and `description` metadata. This workbench uses the spec for skill ingestion and retrieval. It imports markdown instructions/references as data and does not execute imported scripts.

## Migration design implications

- Use an intermediate workflow representation to avoid pairwise source-to-target translators.
- Separate parsing, semantic understanding, target capability mapping, code generation and target-native validation.
- Keep platform/version adapters independent from language parsers.
- Track source evidence, line/procedure mapping, assumptions and unresolved details.
- Never infer credentials, selectors, connection IDs, connector operation schemas or private environment resources.
- Treat third-party skill content as untrusted reference input and retain license/provenance metadata.

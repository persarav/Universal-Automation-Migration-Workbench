---
name: uipath-rpa-migration
description: Analyze, design, generate and validate UiPath Windows XAML projects and workflows for migration from code or other RPA platforms.
---
# UiPath RPA migration skill

## Project-first rules
1. Detect project target framework and project style before generating anything. New or ambiguous projects default to Modern Windows XAML unless the user explicitly requests coded C# or Legacy.
2. Preserve the existing `expressionLanguage` and `targetFramework` when editing a project. Do not silently convert Legacy and Windows projects.
3. Inspect `project.json`, package dependencies, existing workflow conventions and activity availability before authoring XAML.
4. Use a small number of cohesive workflows. Main.xaml orchestrates business phases; child XAML files implement reusable business operations.
5. Prefer native activities for Excel/file/system interactions. Use a small `Invoke Code` or coded workflow for dense table aggregation, scoring and complex transformations rather than deeply nested XAML loops.
6. Use `Invoke Workflow File` with explicit typed arguments for reusable workflows. Avoid global mutable state.

## XAML authoring rules
- Use UiPath Studio-compatible XAML and the correct activity package namespaces.
- Every workflow must have a valid root Activity and appropriate `x:Class` / `x:Members` declarations.
- Give every activity a unique, stable `sap2010:WorkflowViewState.IdRef`.
- Wrap bodies of Sequence, If, ForEach and TryCatch branches in proper Sequence/container nodes as required by the activity schema.
- Use meaningful DisplayName values that describe the action and business purpose.
- Use VB expressions in the correct expression syntax and XML-escape attribute content.
- Define arguments with explicit direction and type. Keep input/output contracts small and documented.
- Do not invent selectors, UI element IDs, connection IDs, tenant IDs or credential names. Mark these as capture/configuration requirements.

## Excel migration
- Prefer Workbook activities for direct workbook operations when Excel UI is not required; use `Read Range Workbook` and `Write Range Workbook` for DataTable-based files.
- For macro-enabled `.xlsm`, preserve the original file unless the process explicitly changes it; never run macros during static analysis.
- Read each required sheet/table once into DataTables, resolve headers by aliases, normalize keys, calculate in memory, and write only changed output columns/ranges.
- Preserve formulas, formatting, hidden sheets and workbook features by writing only targeted cells/columns. Validate write ranges carefully.

## SAP/UI automation
- Prefer SAP GUI activities/official SAP integration where supported. If UI is required, capture selectors in the actual environment using Studio UI Explorer/Object Repository.
- Never fabricate SAP selectors. Store transaction codes, field mappings and selector evidence separately from business-rule logic.
- Add timeout, busy-state checks, status-bar handling, popup recovery and idempotent retry boundaries.

## Error handling and reliability
- Use Try Catch at meaningful transaction boundaries, not around every activity.
- Validate inputs before destructive actions; fail closed on site/DC mismatches or missing deletion indicators.
- Use logs with business keys and safe diagnostics. Never log passwords, tokens or sensitive workbook values unnecessarily.
- Separate business-rule exceptions from system exceptions. Preserve source values unless the corresponding rule authorizes an update.
- Make workflows restartable and record what was changed.

## Validation gates
1. Parse all XAML as XML.
2. Confirm referenced child workflows exist and argument names/types match.
3. Confirm every activity package is present in project.json.
4. Open in the matching UiPath Studio version, resolve package metadata, validate and build.
5. Run regression cases with source and target on the same fixtures and compare outputs and side effects.
6. Capture SAP selectors and validate UI adapters separately.

## Upstream reference
https://github.com/UiPath/skills/tree/main/skills/uipath-rpa

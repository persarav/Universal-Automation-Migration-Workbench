---
name: vba-migration
description: Analyze VBA modules, workbook events, Excel object model usage, procedures, dependencies and business rules for migration into RPA or another language.
---
# VBA migration skill

## Safe source extraction
- Extract `xl/vbaProject.bin` from XLSM/XLSB without executing macros.
- Enumerate modules, class modules, UserForms and document modules; decode compressed VBA source and retain module/line provenance.
- Inventory workbook and worksheet events, references, declared types, constants, public/private scope, and external COM/API dependencies.
- Never run the uploaded macro as part of source analysis.

## Procedure understanding
For every Sub/Function/property/event, capture signature, arguments, return type, local declarations, assignments, loops, branches, error handlers, external calls, workbook/range reads/writes and callees. Build a call graph and distinguish pure helper functions from side-effecting procedures.

## Excel semantics
- Resolve sheet and column names dynamically when source does so; retain header aliases and fallback behavior.
- Distinguish Value, Value2, Text, Formula, UsedRange, CurrentRegion, tables, named ranges, filters and hidden rows.
- Preserve Date serials, Currency/Decimal precision, Empty/Null/Error values, case rules, leading-zero identifiers and locale-specific parsing.
- Record whether a range update changes formulas, comments/notes, formatting, protection, validation or filters.
- Treat Excel event handlers as separate triggers; do not silently convert UI-only layout/protection code into business rules.

## Migration
- Convert business rules to a technology-neutral decision table first.
- Prefer DataTables/dictionaries for bulk row processing in RPA, with reads and writes batched.
- Preserve the exact comment text, order/precedence, blank-cell eligibility and existing-value preservation rules.
- Map SAP/COM/UI calls into capture-required adapters; do not invent selectors or credentials.
- Produce source-to-target mapping for every procedure, including UI-only items and dead/unused routines.

## Regression requirements
Test leading-zero article keys, duplicate keys, blank/error cells, deleted rows, date boundaries, partial/over-supply, existing comments, mismatched sites/DCs, empty exports and all side effects. Compare target comments, dates, workbook data and external system actions against a known-good baseline.

## Upstream reference
https://github.com/RenRMT/claude-code-vba-skills

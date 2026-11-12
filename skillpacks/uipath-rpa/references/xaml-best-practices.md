# UiPath XAML best-practice checklist

- One orchestrator with clear phase boundaries; avoid a single monolithic Main.xaml.
- One child workflow per cohesive business capability, not per tiny helper function.
- Use in/out arguments for DataTables and run context, with explicit types and names.
- Prefer native activities for file access, workbook reads/writes, logs and invoke-workflow orchestration.
- Put complex deterministic transformations into a small, unit-testable Invoke Code block or coded workflow.
- Keep business-rule VB code in a named workflow/region and document input/output contracts.
- Use stable unique workflow/activity IDs and descriptive DisplayName values.
- Use `Try Catch` around stages with meaningful recovery and rethrow behavior.
- No fabricated selectors, credentials, package versions or connector IDs.
- Parse XML, check referenced workflow paths, validate activity packages, then build in Studio.
- Run baseline regression tests before production release.

This is a curated checklist, not a replacement for the current UiPath activity schema or official skill repository.

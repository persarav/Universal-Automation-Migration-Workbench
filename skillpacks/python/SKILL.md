---
name: python-code-migration
description: Understand Python source and translate its behavior into maintainable RPA workflows or other programming languages.
---
# Python code migration skill

## Analysis
Use Python AST parsing rather than regular expressions alone. Inventory modules, imports, classes, functions, decorators, parameters, return annotations, async functions, exceptions, context managers, comprehensions, generators, global state and main guards. Preserve source file and line references.

## Runtime semantics
Record Python version, external packages, environment variables, filesystem/network access, subprocesses, database clients, API calls, timezone behavior, Decimal versus float arithmetic, encoding and exception classes. Identify dynamic behavior (`eval`, reflection, monkey-patching) as a semantic uncertainty.

## Mapping to RPA
Convert triggers to queues/schedules/webhooks, functions to reusable workflows, loops to For Each, branches to If/Flow Decision, exceptions to Try Catch/Scopes, files to supported file activities, and external calls to verified connectors or API activities. Move complex transformations into a small tested code component rather than nested visual activities.

## Output
Generate an importable target artifact only when target schemas and dependencies are known. Otherwise produce a detailed target design with exact function mapping, data contracts, dependency list, error model, test cases and explicitly unresolved platform requirements.

## Safety and tests
Never execute uploaded source during static analysis. Treat package installation and import side effects as untrusted. Create fixtures for nulls, unicode, filesystem paths, dates, numeric precision, API timeouts, retries and failure cases.

## Discovery source
https://github.com/topics/python-skills (topic index; select a concrete repository and inspect its own license/skill format).

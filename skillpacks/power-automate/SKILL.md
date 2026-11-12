---
name: power-automate-migration
description: Map source code and RPA workflows to Power Automate cloud flows or desktop flows with explicit connector, action, expression and environment requirements.
---
# Power Automate migration skill

## Select the correct product
Distinguish cloud flows from desktop flows. Use cloud flows for event, schedule, API and connector orchestration. Use desktop flows for attended/unattended desktop UI tasks. Hybrid designs may invoke desktop flows from cloud flows and require machine/runtime configuration.

## Build from source evidence
1. Identify trigger, input schema, variables, actions, branches, loops, scopes, exception paths, outputs and side effects.
2. Map each source action to a native cloud/desktop action or connector operation.
3. Discover exact connector operation IDs and required parameters from current connector metadata. Never invent operation names or connection references.
4. Use Compose/Initialize variable/Set variable/Condition/Apply to each/Scope/Terminate where semantically appropriate.
5. Preserve data types, null handling, date/time zones, loop order and failure behavior.
6. Build child flows for reusable business capabilities and use typed input/output contracts.

## Expressions
Use Power Automate expression syntax and escape strings correctly. Preserve source truth tables, comparison inclusivity, empty/null semantics, date windows and numeric precision. Do not mechanically translate VB/Python operators without type review.

## Reliability and governance
- Use Scopes with configured run-after policies for try/catch/finally patterns.
- Add retry policies only for transient operations and avoid duplicate non-idempotent actions.
- Use environment variables and connection references; do not hard-code secrets.
- Confirm owner, environment, licensing, DLP policy, connector availability and gateway/machine requirements.
- Validate the exported flow definition and run history in the target tenant before deployment.

## Output contract
Produce flow inventory, trigger/action map, expressions, connector dependency list, connection requirements, exception strategy, test matrix and unresolved configuration list. Label generated flow JSON as a draft until imported and validated in Power Automate.

## Upstream reference
https://github.com/microsoft/power-platform-skills/tree/main/plugins/power-automate

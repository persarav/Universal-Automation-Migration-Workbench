---
name: automation-anywhere-migration
description: Map source code or RPA process exports to Automation Anywhere automation tasks, packages, variables and Bot Runner requirements.
---
# Automation Anywhere migration skill

## Analyze source
Identify bot/task entry points, child bots, commands/packages, variables, credential vault references, queues, work queues, schedules, triggers, exception branches, loops, device/browser interactions and dependencies. Distinguish Automation 360 exports and older versions; inspect manifest/schema and package versions before mapping.

## Target structure
Prefer a small parent bot for orchestration and reusable child bots for business capabilities. Use supported package commands for Excel, files, email, browser, APIs, SAP and queues. Use variables with explicit scope and data types. Model exception handlers, retries, logging and queue item status transitions.

## Platform-specific constraints
Package versions, Bot Agent/Runner setup, Control Room permissions, credential vault, device pools, unattended licenses, schedules and environment settings are not portable from source. Obtain exact target version and package schemas. Do not invent package command names or selectors.

## Deliverables
Produce bot/task hierarchy, source-to-command mapping, variable dictionary, package dependencies, credential/connection inventory, exception strategy, import requirements, regression matrix and unresolved target tenant settings. Treat generated bot packages as drafts until imported and validated in Automation 360.

## Validation
Check package compatibility, credential references, selectors, Control Room permissions, Bot Runner execution, logs, retries and queue behavior in a non-production environment.

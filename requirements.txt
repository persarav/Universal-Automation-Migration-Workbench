---
name: blue-prism-migration
description: Map source code and other RPA exports into Blue Prism Process, Business Object, action, page, data item and release structures.
---
# Blue Prism migration skill

## Analyze source
Extract process/object definitions, pages, stages, data items, collections, environment variables, credentials, business objects, code stages, exception routes, schedules, work queues and release dependencies. Inspect release/package version and XML/schema before generation.

## Target design
Use a Process for orchestration and Business Objects for reusable application-level operations. Split pages by cohesive business capability. Map sequence/decision/loop/multi-computation/action/exception stages, collections and queue operations. Keep environment-specific values and credentials in Blue Prism environment/credential stores.

## Code stages and UI
Use Code Stage only when native stages/objects do not adequately express the transformation. Compile against the target Blue Prism/.NET version. UI selectors, application models, credentials and runtime resource configuration must be captured in the actual environment; never fabricate them.

## Governance and reliability
Preserve queue item statuses, exception types, retry behavior, session lifecycle, audit fields and object initialization/cleanup. Ensure release dependencies and environment variables are defined.

## Deliverables and validation
Create a process/object/page map, data item schema, object/action contracts, code-stage source, exception design, release dependency list and regression plan. Validate import, release, compilation, resource execution and queue outcomes in the target Blue Prism version before declaring it deployable.

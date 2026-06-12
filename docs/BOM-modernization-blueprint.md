# BOM Modernization Blueprint (Portal)

## Scope and constraints
- Source system: BOM.xlsm (Excel macro application with ActiveX, VBA, external connections, shell integrations).
- Analysis mode: read-only only. No database write during reverse engineering.
- Goal: replace BOM with a modern, configurable module inside Gantech Portal.
- Non-goal for phase 1: perfect 1:1 UI clone of Excel visuals.

## What BOM currently does (v1 map)
- Data reads from Visma (22 workbook connections, parameterized by worksheet cells).
- Data orchestration via VBA macros:
  - refresh supplier/customer/product datasets
  - refresh pivots and lookup surfaces
  - run product revision flow
  - export PDF to network path
  - open external workbook for inbox handling
  - run external Python scripts (including DXF flow)
- Data write detected:
  - direct UPDATE on Prod table through ADODB command in macro flow.

Reference artifact: docs/BOM-analysis-v1.json

## Target architecture inside Portal

### 1) Bounded modules
- BOM Core
  - product master lookup, revision input model, component structure.
- Calculators
  - laser, bending, welding, material calculators as separate strategy modules.
- Resources and routing
  - resource registry, route rules, capacity parameters.
- Documents and integrations
  - drawing refs, PDF generation, external script adapter abstraction.

### 2) Read path (cache-first)
- Query Adapter Layer
  - each query is defined as a versioned template with typed parameters.
- Cache layers
  - L1 in-memory cache for hot requests.
  - L2 persistent cache (disk/Redis) with TTL and tags.
- Invalidation
  - time-based TTL + manual invalidation + domain-event invalidation.
- Request coalescing
  - in-flight deduplication per cache key.

### 3) Write path (controlled)
- Explicit command endpoints only (no hidden write side effects).
- All writes behind feature flags and role checks.
- Audit trail: who, when, before/after payload, target record.
- Dry-run mode for write commands in acceptance tests.

### 4) Query configurability model
- Table: query_definitions
  - key, version, sql_template, param_schema, default_ttl, enabled.
- Table: query_bindings
  - module_name, action_name, query_key, transform_name.
- Table: tenant_query_overrides (optional)
  - tenant_id, query_key, version pin/override.
- SQL safety
  - allow only parameter placeholders from schema.
  - deny ad-hoc string concatenation.

### 5) Customization model for future changes
- resource_types and resources configurable in DB.
- rule_sets for calculators and routing, versioned.
- ui_layout_state per user (drag/drop and panel arrangement).
- mapping tables for Visma field aliases and fallback logic.

## Proposed BOM-to-Portal flow mapping
- Excel "refresh macro" -> Portal "refresh endpoint + background cache warm".
- Excel pivot refresh -> Portal materialized read model + chart API.
- Excel product revision macro -> Portal revision wizard + controlled command.
- Excel PDF export -> server-side PDF service (template based).
- Excel external workbook open -> inbox module API + queue processor.
- Excel shell/python -> integration worker service with secure task queue.

## Performance plan (Visma load reduction)
- Replace workbook-wide RefreshAll with scoped reads by module/action.
- Cache TTL defaults:
  - static dimensions: 8-24h
  - product lookup: 2-8h
  - revision candidate data: 30-120m
  - capacity snapshots: 5-30m
- Add query observability:
  - latency, rows, cache hit rate, error classes, top expensive query keys.
- Add rate limiting / concurrency caps on heavy endpoints.

## Security and reliability
- No direct DSN from client; backend-only DB access.
- Secrets in environment or secret manager only.
- Least-privilege DB credentials:
  - read role for query adapters
  - separate write role for command endpoints.
- Integration sandbox for script execution.

## Migration phases

### Phase 0 - reverse engineering hardening
- Complete macro flow map (trigger, inputs, outputs, side effects).
- Query catalog extraction (all workbook connections + parameters).
- External dependency inventory (paths, files, scripts, credentials, DSN assumptions).

### Phase 1 - read-only parity
- Implement BOM Read API in portal.
- Add cache and observability.
- Reproduce critical read screens and filters.

### Phase 2 - controlled write parity
- Implement revision update flow as explicit command.
- Add approval/audit and dry-run mode.
- Add PDF/document pipeline.

### Phase 3 - decommission Excel dependencies
- Replace shell/python direct calls with worker tasks.
- Retire external workbook coupling.
- Keep legacy fallback only behind emergency switch.

## Acceptance criteria (initial)
- Functional parity on top 10 BOM business flows.
- No silent DB writes.
- 70%+ cache hit on read-heavy endpoints under normal usage.
- Measurable reduction of peak Visma DB load compared with current workbook usage.

## Open questions to close next
- Exact authoritative write targets beyond Prod update.
- Ownership and lifecycle of external workbook "Indhold af indbakke.xlsm".
- Required behavior of Python scripts and expected outputs.
- Role matrix for who can run revisions/exports/writes.

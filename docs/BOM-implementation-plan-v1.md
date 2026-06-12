# BOM Implementation Plan v1 (Step-by-step)

## Principles
- Start read-only parity first.
- No database write until write contract is explicitly approved.
- Keep Excel BOM available as fallback during phased rollout.

## Phase 0 - Discovery hardening (current)
- Deliverables
  - BOM analysis JSON
  - Flow map
  - Query catalog
  - BOM->Portal mapping
- Exit criteria
  - Top flows and dependencies are documented and reviewed.

## Phase 1 - Read API foundation
- Scope
  - Build /bom read endpoints for supplier, customer, product, resources, materials.
  - Implement query adapter with typed parameters.
  - Implement cache layer (L1 memory + persistent cache).
- Technical tasks
  - Add bomQueryAdapter service.
  - Add query registry config table (versioned templates).
  - Add cache tagging and invalidate endpoint.
  - Add observability (latency, row count, cache hit/miss).
- Exit criteria
  - Read endpoints replace the most used lookup macros.
  - DB load reduced vs manual Excel refresh cycle.

## Phase 2 - Stykliste parity UI (read)
- Scope
  - Build Stykliste views in portal using read APIs.
  - Reproduce filtering and key lookup interactions.
- Technical tasks
  - Implement Stykliste page module.
  - Add customer/product drawing lookup panel.
  - Add resource/material side panels from cached APIs.
- Exit criteria
  - Users can execute main Stykliste read workflows without Excel.

## Phase 3 - Revision command (controlled write)
- Scope
  - Replace Opd_Prod_vn with explicit command endpoint.
- Technical tasks
  - Define revision DTO schema.
  - Implement POST /bom/revisions/apply with dry-run mode.
  - Add audit table and role guard.
  - Add post-write invalidation for affected cache keys.
- Exit criteria
  - Controlled write path validated against expected Prod updates.

## Phase 4 - Document and integration replacement
- Scope
  - Replace GemSomPDF and shell/python dependencies.
- Technical tasks
  - Implement document generation service.
  - Implement integration worker queue for DXF/token tasks.
  - Remove direct shell execution from user session path.
- Exit criteria
  - No operational dependency on local C:\SCRIPTS or external workbook open.

## Phase 5 - Cutover and de-risk
- Scope
  - Pilot users run Portal BOM as primary.
  - Excel kept as emergency fallback only.
- Technical tasks
  - Feature flags per team/user.
  - Compare output parity on sample orders/revisions.
  - Lock and monitor write commands.
- Exit criteria
  - Portal BOM stable, monitored, and accepted.

## Initial backlog (ordered)
1. Query adapter and parameter schema
2. Query catalog registry + versioning
3. Cache service + tag invalidation
4. Supplier/customer/product/resource/material endpoints
5. Stykliste read UI module
6. Revision DTO + dry-run command
7. Audit logging for write commands
8. PDF generation service
9. Integration worker for DXF/token
10. Feature-flagged rollout and telemetry dashboard

## KPIs
- P95 response time for read endpoints
- Cache hit ratio per endpoint
- Number of DB queries per user action
- Write command success/failure and rollback rate
- User adoption rate vs Excel fallback usage

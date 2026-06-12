# BOM to Portal Mapping v1

## Objective
Map BOM Excel flows to explicit Portal services, endpoints, and cache keys.

## Naming conventions
- Endpoint prefix: /bom
- Service prefix: bom*
- Cache key prefix: bom:v1:

## Flow mapping matrix

| BOM flow | Excel trigger | Portal endpoint | Service | Cache keys | Write? |
|---|---|---|---|---|---|
| Supplier refresh | Knap248_Klik / Knap251_Klik | GET /bom/suppliers | bomSupplierService | bom:v1:suppliers:list | No |
| Supplier pivot summary | same flow | GET /bom/suppliers/summary | bomSupplierSummaryService | bom:v1:suppliers:summary | No |
| Customer lookup | Kunder/Visma lookup tables | GET /bom/customers | bomCustomerService | bom:v1:customers:list | No |
| Product lookup by customer | VISMA_EXCEL parameterized query | GET /bom/products?customerNo= | bomProductService | bom:v1:products:customer:{customerNo} | No |
| Drawing/revision lookup by TgNo | Visma6 + Visma7 queries | GET /bom/revisions/by-drawing?tgn=...&cust=... | bomRevisionLookupService | bom:v1:revisions:drawing:{cust}:{tgn} | No |
| Resource/routing data | Ressourcer + FreeInf tables | GET /bom/resources | bomResourceService | bom:v1:resources:list | No |
| Material stock and pricing | råvarer + RV-beholdn | GET /bom/materials | bomMaterialService | bom:v1:materials:list:{filterHash} | No |
| Laser parameters | skæreparametre / Laserberegner | GET /bom/calculators/laser-params | bomLaserParamService | bom:v1:laser:params | No |
| Bending calculator data | Bukkeberegner query | GET /bom/calculators/bending-params | bomBendingParamService | bom:v1:bending:params | No |
| Start revision context | GåTil_NyRev | POST /bom/revision-sessions | bomRevisionSessionService | bom:v1:revision-session:{sessionId} | No |
| Apply revision to product | Opd_Prod_vn (UPDATE Prod) | POST /bom/revisions/apply | bomRevisionCommandService | invalidate revision/product keys | Yes (controlled) |
| PDF export | GemSomPDF | POST /bom/documents/revision-pdf | bomDocumentService | bom:v1:doc-template:{templateId} | File write |
| Inbox bridge | Åben_tjek_indbakke / NyRev_Indbakke | GET /bom/inbox/items | bomInboxService | bom:v1:inbox:list | No |
| API token helper | EseguiScriptPython | POST /bom/integrations/token/refresh | bomIntegrationService | bom:v1:token:last | No DB |
| DXF helper | EseguiScriptPythonConSelezioneDXF | POST /bom/integrations/dxf/convert | bomDxfService | bom:v1:dxf:{fileHash} | No DB |

## Query ownership by service
- bomSupplierService
  - connection 1 (Actor suppliers)
- bomProductService
  - connections 2, 8, 19, 20 (Prod lookups)
- bomResourceService
  - connections 3, 4, 10, 21
- bomMaterialService
  - connections 5, 6
- bomCustomerService
  - connections 9, 11, 12, 13
- bomDocumentService
  - connection 14 + template/render layer
- bomMetaService
  - connections 16, 18, 22

## Cache policy v1
- Dimensions (Txt, users, static lookup): TTL 24h
- Master lists (suppliers/customers/resources): TTL 8h
- Product lookups by customer: TTL 2h
- Revision lookup by drawing/customer: TTL 30m
- Material/stock snapshots: TTL 15m
- Calculator params: TTL 8h

## Invalidation policy
- Manual invalidate endpoint: POST /bom/cache/invalidate
- Scoped invalidation tags:
  - suppliers
  - customers
  - products:{cust}
  - revisions:{cust}:{tgn}
  - materials
  - calculators
- Automatic invalidation on write command /revisions/apply

## Write safety rules (mandatory)
- No SQL write from UI layer.
- Writes only through command service with DTO validation.
- Every write must produce an audit record.
- Dry-run mode required for each write command.
- Feature flag gate: bom.write.enabled

# BOM Flow Map v1 (Read-Only Reverse Engineering)

## Method
- Source inspected: BOM.xlsm package structure + extracted VBA dump.
- Analysis mode: static read-only only.
- No macro execution, no DB write by analysis tooling.

## High-level domains
- Master data and lookup
  - sheets: Visma FV, Kunder, Lev, LevPiv, TgNo, Usr, Ressourcer.
- BOM and revision workflow
  - sheets: Stk.liste opr., Ny rev, Ny revision, NyRevision, Stykliste.
- Calculators
  - sheets: Laserberegner, Laserberegner2, Bukkeberegner, Svejseberegner, MatrBeregn, råvarer, skæreparametre.
- Integration/documents
  - sheets: API WEB, U-lev, U-lev Doc.

## Identified executable procedures (non-empty)
- Knap248_Klik
- Knap251_Klik
- GemSomPDF
- Opd_Prod_vn
- NyRev_omdøb_og_flyt
- NyRev_Indbakke
- GåTil_NyRev
- ReturTilForside
- Åben_tjek_indbakke
- Tegnnr_i_Visma
- TomFelt
- Få_tastefelter
- Mange_Tastefelter
- sj
- EseguiScriptPythonConSelezioneDXF
- LeggiTokenDagliAppunti
- EseguiScriptPython

## Flow catalog (v1)

### Flow A: Supplier/Pivot refresh
- Trigger
  - macro Knap248_Klik / Knap251_Klik (button style action).
- Inputs
  - active sheet/list object selection in Lev.
- Steps
  - QueryTable.Refresh on Lev table.
  - PivotCache.Refresh on LevPiv.
  - return to Stk.liste opr.
- Outputs
  - refreshed supplier-related table/pivot state.
- Side effects
  - DB read load via workbook connections.

### Flow B: Start revision context
- Trigger
  - macro GåTil_NyRev.
- Inputs
  - current user environment variable USERNAME.
- Steps
  - writes username into worksheet cell.
  - navigates to NyRevision and prepares range formulas/selection.
- Outputs
  - revision prep context in workbook cells.
- Side effects
  - local workbook state mutation only.

### Flow C: Product revision apply (critical write)
- Trigger
  - macro Opd_Prod_vn.
- Inputs
  - rows from Ny revision sheet.
  - connection string from Opsætning sheet cell.
- Steps
  - copy staging data from Ny rev to Ny revision.
  - iterate rows marked for update.
  - build and execute ADODB UPDATE Prod statements.
  - mark updated rows in sheet.
  - clear staging range.
  - run GemSomPDF.
  - ActiveWorkbook.RefreshAll.
- Outputs
  - DB product fields updated (Inf7, Inf8, Inf2, PictNo).
  - workbook refreshed.
  - PDF exported.
- Side effects
  - direct DB write.
  - file system write to network path.

### Flow D: PDF export
- Trigger
  - macro GemSomPDF (standalone and called by Opd_Prod_vn).
- Inputs
  - active sheet content and filename cell.
  - archive path P:\Visma\Stykliste\Arkiv\.
- Steps
  - set print/page layout.
  - ensure target directory exists (MkDir).
  - ExportAsFixedFormat to PDF.
- Outputs
  - PDF file in archive path.
- Side effects
  - file system write on network share.

### Flow E: Inbox workbook bridge
- Trigger
  - macro Åben_tjek_indbakke and related NyRev_Indbakke flow.
- Inputs
  - external workbook path/name Indhold af indbakke.xlsm.
- Steps
  - Workbooks.Open external workbook.
- Outputs
  - external workbook activated.
- Side effects
  - hard dependency on sibling workbook.

### Flow F: Visma drawing/product lookup refresh
- Trigger
  - macro Tegnnr_i_Visma.
- Inputs
  - table query and pivot on Visma FV sheet.
- Steps
  - QueryTable.Refresh on Tabel_Forespørgsel_fra_VISMA_EXCEL.
  - PivotCache.Refresh.
  - copy helper ranges on NyRevision.
- Outputs
  - refreshed lookup data and helper rows for revision.
- Side effects
  - DB read load.

### Flow G: External Python automation (API token)
- Trigger
  - macro EseguiScriptPython.
- Inputs
  - script path C:\SCRIPTS\Nexting.py.
- Steps
  - Shell executes python script.
  - wait 4 seconds.
  - read clipboard token.
  - write token into API WEB!Z1.
- Outputs
  - token imported to workbook.
- Side effects
  - shell execution + clipboard dependency.

### Flow H: External Python automation (DXF)
- Trigger
  - macro EseguiScriptPythonConSelezioneDXF.
- Inputs
  - selected DXF file via FileDialog.
  - script path C:\SCRIPTS\DXF2JSON.py.
- Steps
  - user selects DXF.
  - Shell executes python script with DXF path.
  - waits for completion window.
- Outputs
  - expected transformed artifact (outside workbook, inferred).
- Side effects
  - shell execution + local script path dependency.

## Data access inventory (from workbook connections)
- Approx. 22 workbook connections.
- Main read tables detected
  - Actor
  - Prod
  - PrDcMat
  - BgtLn
  - Txt
  - StcBal
  - FreeInf1
  - FreeInf2
  - ActInf
  - AssLink
  - R8
  - Doc
  - DocLink

## Risk points for migration
- Hidden business logic in worksheet formulas and named ranges.
- Direct SQL write in VBA without explicit API boundaries.
- Coupling to network paths and external workbook.
- Shell/Python integration not versioned nor centrally governed.

## Migration priority (suggested)
1. Read-only parity for top lookup/refresh flows.
2. Controlled revision write workflow with audit and dry-run.
3. Replace external workbook dependency with inbox module.
4. Replace shell-based scripts with managed worker jobs.

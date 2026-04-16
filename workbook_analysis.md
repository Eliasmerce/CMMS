# Excel Workbook Functional Analysis (Inferred from Provided Structure)

## Scope and assumption
This assessment is based on the user's structural description of the `.xlsm` workbook because no workbook file is present in this repository for direct inspection.

## 1) Purpose of each major sheet

### Formulario
Primary user interface for front-line maintenance reporting.
- Captures report metadata (report number, date, personnel, equipment).
- Captures intervention details (materials, quantities, description).
- Captures downtime period (start/end and total downtime).
- Serves as controlled input point before persistence.

### Base de Datos
Historical transactional table (one record per report).
- Stores normalized row-wise records for later aggregation.
- Acts as source of truth for indicators, pivots, and charts.
- Should preserve immutable historical entries.

### Calculo Indicadores
KPI engine layer.
- Transforms historical log into operational metrics.
- Typical metrics: unplanned downtime, cost impact, production loss, target vs actual availability.
- Converts raw event logs into management-level indicators.

### Summary/chart sheets
Decision-support and visibility layer.
- Aggregate failures by area/equipment.
- Show availability trends and critical equipment performance.
- Present trend charts for operational review.

### Support sheets
Data model and control layer.
- Dropdown sources and catalog tables (equipment, areas, personnel, materials).
- Helper calculations and staging ranges for summaries/charts.
- Should remain hidden/protected to reduce user error.

## 2) Data flow (entry -> storage -> KPIs -> visualization)
1. **User entry in Formulario**: operator enters work report and downtime fields.
2. **Validation and save macro**: VBA validates mandatory fields, computes/normalizes time fields, and appends a row.
3. **Append to Base de Datos**: record is written as an immutable transaction row.
4. **KPI recalculation**: formulas in Calculo Indicadores read Base de Datos and compute availability/cost/loss metrics.
5. **Summary refresh**: pivot-like or formula summaries aggregate by area/equipment/time period.
6. **Chart update**: charts consume summary ranges and display trends/availability/failure counts.
7. **Optional reverse lookup**: Formulario may retrieve prior report data for correction/reprint (if implemented).

## 3) Strengths observed in this architecture
- Clear separation between UI (Formulario), storage (Base de Datos), and analytics (Calculo Indicadores).
- Low training barrier for operations teams familiar with Excel.
- Fast implementation and adaptation without enterprise CMMS deployment.
- Macro-enabled workflow can enforce process consistency.

## 4) Weaknesses / risk points

### Usability risks
- Form complexity can cause incomplete or inconsistent entries.
- Manual start/end time handling may create negative or cross-day errors.
- Free-text fields may reduce downstream analysis quality.

### Structural/data-quality risks
- If Base de Datos is not an Excel Table, range drift can break formulas/charts.
- Duplicate report IDs without uniqueness checks.
- Mixed data types (text dates, numeric stored as text) can corrupt KPIs.

### Automation/VBA risks
- Macro-disabled environment breaks critical save/refresh actions.
- Unhandled VBA errors can partially write records.
- Hard-coded ranges/sheet names make workbook fragile to edits.

### Compatibility risks
- `XLOOKUP` requires modern Excel versions; older perpetual versions may fail.
- Dynamic arrays may not spill correctly in legacy versions.
- Locale differences (date/time separators, decimal commas) can alter calculations.

### Control/audit risks
- No audit trail for edits/deletes in Base de Datos.
- Weak sheet protection enables accidental formula overwrite.
- Multi-user simultaneous access can cause conflicts or lost updates.

## 5) Practical improvements (without changing business purpose)

### A. Usability
- Convert Formulario fields into a guided sequence with visual required markers.
- Add immediate validation prompts (missing mandatory fields, invalid time spans).
- Add controlled lists for equipment, failure mode, area, and personnel.

### B. Data structure
- Convert Base de Datos to an official Excel Table (`tblReportes`) and reference structured columns.
- Add stable surrogate key (auto ID) plus business key (report number).
- Standardize datetime columns: `StartDateTime`, `EndDateTime`, `DowntimeHours` numeric.

### C. KPI robustness
- Centralize constants (cost/hour, target availability) in a parameter table.
- Replace volatile formulas and wide whole-column references with bounded table formulas.
- Add data quality KPIs (missing end time, zero downtime, duplicate report number).

### D. VBA hardening
- Add explicit `Option Explicit` and centralized error handling/logging.
- Use transactional write pattern: validate -> write row -> verify -> timestamp -> success message.
- Replace hard-coded addresses with named ranges/table references.

### E. Compatibility strategy
- For mixed-version environments, provide `INDEX/MATCH` fallback where `XLOOKUP` is used.
- Document minimum supported Excel version (e.g., Microsoft 365 / Excel 2021+).
- Add startup compatibility check macro with user-facing warning.

### F. Governance and reliability
- Protect formula/KPI/support sheets; unlock only input cells in Formulario.
- Add backup/archive macro (daily copy with timestamp).
- Introduce change log sheet for macro version, formula changes, and data corrections.

## 6) Recommended target operating model
- **Operators**: only interact with Formulario.
- **Supervisors**: consume summaries and validate exceptions.
- **Power users/admin**: maintain support catalogs and parameters.
- **Workbook owner**: controls macro updates/versioning and periodic integrity checks.

## 7) Quick validation checklist (for the actual file)
- Save macro blocks incomplete forms.
- Duplicate report number is rejected.
- Cross-midnight downtime is computed correctly.
- Base de Datos appends exactly one row per save.
- KPIs reconcile with manual sample calculations.
- Charts refresh after new record insertion.
- Workbook opens with clear warning if macros are disabled.

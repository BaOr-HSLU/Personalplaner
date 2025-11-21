# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Personalplaner** is an Excel VBA-based personnel planning and resource management system. The repository contains exported VBA modules from an Excel workbook (.xlsm) used for managing employee resources, absences, and workload planning.

**Language**: Bilingual - Code/comments in English, UI messages in German
**Platform**: Microsoft Excel with VBA (VBA7 and legacy VBA compatible)
**Current Version**: v2.8
**Architecture**: Service-oriented with domain models and helper utilities

## Repository Structure

This repository contains **exported VBA modules** from an Excel workbook, not a standalone application. File types:

- **`.bas`** - VBA Standard Modules (services, helpers, calculations)
- **`.cls`** - VBA Class Modules (domain models with PredeclaredId)
- **`.frm`** - VBA UserForms (UI dialogs)
- **`.frx`** - UserForm binary data (compiled form resources)
- **`.doccls`** - Document Class Modules (worksheet/workbook event handlers)

### Architecture (v2.8)

v2.8 introduces a **service-oriented architecture** with clear separation of concerns:

#### Domain Models (`*.cls`)
| Class | Purpose |
|-------|---------|
| `Employee.cls` | Employee entity with contact info (Name, Email, Phone, Function, Team) |
| `Project.cls` | Project entity with commission numbers and remarks |
| `AbsenceCode.cls` | Absence code definitions with short/long forms and colors (PredeclaredId) |

#### Service Modules (`*Service.bas`)
| Service | Purpose |
|---------|---------|
| `CalendarService.bas` | Calendar creation with work days (Mon-Fri), KW, holidays, vacations |
| `WorkloadCalculations.bas` | **UDFs** for workload calculations, optimized for 200+ employees over 5 years |
| `EmployeeService.bas` | Employee data management, loading from sheets |
| `ProjectService.bas` | Project CRUD operations on project master sheet |
| `WeeklyReportService.bas` | Weekly report generation and PDF export with email |
| `WeeklySheetService.bas` | Weekly plan sheet creation from templates |
| `FilterService.bas` | Filtering functionality for employee/project views |
| `EmailService.bas` | Email composition and sending via Outlook integration |

#### Helper Modules (`*Helpers.bas`)
| Helper | Purpose |
|--------|---------|
| `DateHelpers.bas` | Date/calendar week utilities, robust date column finding |
| `ValidationHelpers.bas` | Data validation routines, input checks |

#### UI Controller
| Module | Purpose |
|--------|---------|
| `RibbonController.bas` | Custom Ribbon UI management (IRibbonUI callbacks) |

#### UserForms
| Form | Purpose |
|------|---------|
| `UF_Filter.frm` | Filter dialog for employee views |
| `UF_Projekte.frm` | Project management dialog |
| `UF_ProjektErstellen.frm` | Project creation dialog |

#### Worksheet Classes
| Class | Purpose |
|-------|---------|
| `DieseArbeitsmappe.doccls` | Workbook-level event handlers |
| `Tabelle1.doccls` | Main planning sheet (Tabelle1 CodeName) |
| `Tabelle3.doccls` | Auslastung/workload sheet (Tabelle3 CodeName) |
| `Tabelle4.doccls`, `Tabelle7.doccls`, etc. | Other worksheet handlers |
| `Diagramm1.doccls` | Chart sheet handler |
| `shWRTemplate.doccls` | Weekly report template sheet |

## Key Technical Concepts

### 1. Exported VBA Files
The `.bas`, `.frm`, and `.doccls` files are **text exports** of VBA code from the Excel workbook. To use them:
- Import into an Excel workbook via VBA Editor
- Or maintain as version-controlled source files and re-import when needed

### 2. Custom Ribbon UI
The system uses a **Custom Ribbon** interface (IRibbonUI) defined in `CustomUI.bas`:
- Provides buttons for navigation (Heute, Übersicht, Auswertung, Filter, Projekt)
- Context-sensitive elements for different worksheet types
- Ribbon updates require Excel restart after changes

### 3. Data Structure
- **ListObject-based** tables for structured data management
- Tables to maintain: `Feiertage` (holidays), `Ferien` (school vacations), `Mitarbeiter` (employees)
- Calendar worksheets with merged cells for KW/month/year headers

### 4. Absence Codes
Standard codes used throughout the system:
- **F** = Ferien (vacation)
- **Fx** = Ferien nicht bewilligt (vacation not approved)
- **K** = Krank (sick)
- **U** = Unfall (accident)
- **WK** = Militär (military service)
- **S** = Schule (school)
- **ÜK** = Überbetrieblicher Kurs (inter-company course)
- **T** = Teilzeit (part-time)

### 5. UDFs (User Defined Functions)
Key Excel functions defined in `WorkloadCalculations.bas`:
```vba
GetWorkloadByDate(targetDate, [columnOffset], [headerRowNumber], [dataStartRowNumber], [anchorColumnNumber])
  ' Returns workload value for a specific date with optional column offset
  ' Optimized for 200+ employees over 5 years

CountAbsentEmployees(targetDate)
  ' Counts absent employees on a given date based on absence codes

CountAbsenceCodes(targetRange)
  ' Counts cells containing absence codes (F, U, K, WK, S, ÜK, T)

CalculateWorkloadWithExclusion(exclusionRange, [departmentFilter])
  ' Calculates utilization rate excluding specific criteria

CountAvailableEmployees(exclusionRange, [departmentFilter])
  ' Counts available employees for a given day
```

Helper functions in `DateHelpers.bas`:
```vba
FindDateColumn(targetSheet, headerRowNumber, searchDate, [firstColumnToSearch], [lastColumnToSearch])
  ' Robust date column finder (handles date formats, text dates, dates with time)
```

### 6. Performance Optimizations
- Calculation mode set to **Manual** (use F9 or Ribbon "Berechnen" to recalculate)
- `Application.ScreenUpdating = False` during intensive operations
- Dictionary-based lookups instead of loops

## Development Workflow

### Working with VBA Files

When making changes to VBA code:

1. **Editing exported files**: Modify `.bas`/`.frm`/`.doccls` files directly in repository
2. **Testing changes**: Import modified files into Excel workbook via VBA Editor
3. **Version control**: Commit exported VBA files (text-based, diff-friendly)

### Re-exporting from Excel

If the Excel workbook is modified:
```vba
' Export all modules programmatically
Sub ExportAllModules()
    Dim vbComp As VBComponent
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        vbComp.Export "path\to\repo\" & vbComp.Name & GetExtension(vbComp.Type)
    Next
End Sub
```

### Adding New Features

1. **Calendar functions**: Modify `CalendarService.bas` (e.g., add new calendar types, date ranges)
2. **Calculations/UDFs**: Add to `WorkloadCalculations.bas`
3. **Employee operations**: Modify `EmployeeService.bas` or `Employee.cls`
4. **Project operations**: Modify `ProjectService.bas` or `Project.cls`
5. **UI changes**: Update `RibbonController.bas` for Ribbon, or create/modify `.frm` files for dialogs
6. **New worksheets**: Create corresponding `.doccls` file with event handlers
7. **Utility functions**: Add to `DateHelpers.bas` or `ValidationHelpers.bas` as appropriate

### Ribbon Customization

The Custom Ribbon is managed by `RibbonController.bas`:
- XML callbacks defined in CustomUI XML must match function names in `RibbonController.bas`
- Key callbacks: `OnLoad_PersonalPlaner`, `GetControlVisibility`
- Ribbon pointer stored in named range "RibbonID" for dynamic updates
- Control IDs in code must match XML (e.g., "TabDashboard", "TabWeeklyPlan")
- Changes to Ribbon XML require Excel restart to take effect
- Use `RefreshRibbon()` to invalidate and reload ribbon programmatically

## Common Patterns

### Date Handling
The system uses robust date detection via `DateHelpers.bas`:
- `FindDateColumn()` handles Date values, text dates (dd.mm.yyyy), and dates with time components
- Always uses `Int(CDbl(dateValue))` to strip time portions for comparisons
- Multiple search strategies: direct MATCH on serial numbers, then text format matching, then cell-by-cell scan
- Header rows contain dates; use `FindDateColumn()` to locate columns

### Error Handling
Functions return Excel error values:
- `CVErr(xlErrNA)` - #NV (not found)
- `CVErr(xlErrValue)` - #WERT! (invalid input)
- `CVErr(xlErrRef)` - #BEZUG! (reference error)

### ListObject Access
Use structured references:
```vba
Dim tbl As ListObject
Set tbl = ws.ListObjects("Mitarbeiter")
tbl.DataBodyRange.Rows(1).Cells(1).Value
```

### Event Management
To prevent recursion during worksheet events:
```vba
Application.EnableEvents = False
' ... make changes ...
Application.EnableEvents = True
```

### Domain Models (Class Modules)
v2.8 introduces proper domain models:

**Employee.cls**:
```vba
Dim emp As Employee
Set emp = New Employee
emp.DisplayName = "Max Mustermann"
emp.EmailAddress = "max@example.com"
emp.ParseFromCellValue(cellValue)  ' Parse from multi-line cell
If emp.HasValidEmail() Then ' validation
```

**Project.cls**:
```vba
Dim proj As Project
Set proj = New Project
proj.ProjectName = "Project A"
proj.CommissionNumber = "12345"
```

**AbsenceCode.cls** (PredeclaredId):
```vba
' Static factory pattern with PredeclaredId
Dim allCodes As Dictionary
Set allCodes = AbsenceCode.GetAllCodes()

Dim codesArray As Variant
codesArray = AbsenceCode.GetCodesArray(useShortForm:=True)

If AbsenceCode.IsValidCode("F") Then ' validation
```

### Service Pattern
Services encapsulate business logic and data access:

**CalendarService.bas**:
```vba
Call CalendarService.CreateWorkDayCalendar(startCell)
Call CalendarService.InsertHolidaysAndVacations()
```

**EmployeeService.bas**:
```vba
Dim employees As Collection
Set employees = EmployeeService.LoadAllEmployees()
```

**ProjectService.bas**:
```vba
Dim proj As Project
Set proj = ProjectService.LoadProject("Project A")
If ProjectService.SaveProject(proj) Then
```

## v2.8 Key Improvements

v2.8 represents a major architectural refactoring:

### Code Quality
- **English module names** for better maintainability
- **@Folder annotations** for logical organization (Services, Domain, UI, Utilities)
- **@ModuleDescription** for clear documentation
- **Service-oriented architecture** with separation of concerns
- **Domain models** as proper class modules
- **Helper utilities** extracted into dedicated modules

### Performance
- Optimized for **200+ employees over 5 years**
- `WorkloadCalculations.bas` uses efficient lookup strategies
- Date finding with multiple strategies (MATCH, text format, scan)
- CodeName references for faster sheet access (e.g., `Tabelle3`)

### Maintainability
- Clear separation: Services, Models, Helpers, UI
- Single Responsibility Principle per module
- Reusable helper functions (`DateHelpers`, `ValidationHelpers`)
- PredeclaredId pattern for static factory methods (`AbsenceCode`)

### Migration from v2.7
| v2.7 Module | v2.8 Module |
|-------------|-------------|
| `mKalender.bas` | `CalendarService.bas` |
| `mBerechnung.bas` | `WorkloadCalculations.bas` |
| `mAuslastung.bas` | Merged into `WorkloadCalculations.bas` |
| `mKWBlatt.bas` | `WeeklySheetService.bas` + `WeeklyReportService.bas` |
| `mFilter.bas` | `FilterService.bas` |
| `CustomUI.bas` | `RibbonController.bas` |
| (none) | `Employee.cls`, `Project.cls`, `AbsenceCode.cls` (new) |
| (none) | `DateHelpers.bas`, `ValidationHelpers.bas` (new) |
| (none) | `EmployeeService.bas`, `ProjectService.bas`, `EmailService.bas` (new) |

## Version Conventions

- Major versions (v2.x, v3.x): Significant feature additions or architectural changes
- Branch naming: `claude/description-vX.X-sessionID`
- Commits: English, descriptive of changes
- Release notes: Maintained in separate `RELEASE_NOTES_vX.X.md` file (if needed)

## Important Notes

- **Excel-Specific**: This code only runs within Excel VBA environment
- **Bilingual**: Code/comments in English (v2.8), UI messages still in German
- **No Standalone Tests**: Testing requires manual Excel workbook testing
- **Binary Files**: `.frx` files are binary and not human-readable
- **Worksheet CodeNames**: References like `Tabelle3` are CodeNames (not visible sheet names)
- **Manual Calculation**: Always consider recalculation implications when modifying UDFs
- **Dictionary Dependency**: Code uses `Scripting.Dictionary` - ensure Microsoft Scripting Runtime reference
- **VBA7 Compatibility**: Uses conditional compilation for 32-bit/64-bit compatibility
- **Ribbon XML**: Ribbon control IDs in `RibbonController.bas` must match CustomUI XML
- **PredeclaredId**: `AbsenceCode.cls` uses PredeclaredId for static factory pattern

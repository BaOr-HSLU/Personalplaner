# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Personalplaner** is an Excel VBA-based personnel planning and resource management system. The repository contains exported VBA modules from an Excel workbook (.xlsm) used for managing employee resources, absences, and workload planning.

**Language**: German (code comments, documentation, and UI are in German)
**Platform**: Microsoft Excel with VBA (VBA7 and legacy VBA compatible)
**Current Version**: v2.7

## Repository Structure

This repository contains **exported VBA modules** from an Excel workbook, not a standalone application. File types:

- **`.bas`** - VBA Standard Modules (code modules with public/private procedures and functions)
- **`.frm`** - VBA UserForms (UI dialogs)
- **`.frx`** - UserForm binary data (compiled form resources)
- **`.doccls`** - Document Class Modules (worksheet/workbook event handlers)

### Core Modules

| Module | Purpose |
|--------|---------|
| `mKalender.bas` | Calendar creation with work days (Mon-Fri), KW (calendar weeks), holidays |
| `mBerechnung.bas` | **UDFs** for workload calculations, date lookups, availability calculations |
| `mAuslastung.bas` | Additional workload/utilization functions |
| `mKWBlatt.bas` | Weekly plan sheet creation and PDF export |
| `mFilter.bas` | Filtering functionality for employee views |
| `mFormatierung.bas` | Conditional formatting and visual styling |
| `mWertesammler.bas` | Data collection and aggregation utilities |
| `mDatenüberprüfung.bas` | Data validation routines |
| `CustomUI.bas` | Custom Ribbon UI integration (IRibbonUI) |
| `DieseArbeitsmappe.doccls` | Workbook-level event handlers |
| `UF_Filter.frm` | Filter dialog UserForm |
| `UF_Projekte.frm` | Project management UserForm |
| `UF_ProjektErstellen.frm` | Project creation UserForm |
| `shWRTemplate.doccls` | Worksheet template class |
| `Tabelle*.doccls` | Worksheet-specific event handlers and logic |

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
Key Excel functions defined in `mBerechnung.bas`:
```vba
VerweisMABAuslastungTotal(Datum, [offset])
  ' Returns workload value for a specific date with optional column offset

AuslastungMitAusschluss(rngAusschluss, [abteilung])
  ' Calculates utilization rate excluding specific criteria

VerfuegbareMitarbeiter(rngAusschluss, [abteilung])
  ' Counts available employees for a given day

FindeDatumsspalte(ws, HeaderRow, Suchdatum)
  ' Robust date column finder (handles date/text formats)
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

1. **Calendar functions**: Modify `mKalender.bas` (e.g., add new calendar types, date ranges)
2. **Calculations/UDFs**: Add to `mBerechnung.bas` or `mAuslastung.bas`
3. **UI changes**: Update `CustomUI.bas` for Ribbon, or create/modify `.frm` files for dialogs
4. **New worksheets**: Create corresponding `.doccls` file with event handlers

### Ribbon Customization

The Custom Ribbon is defined using IRibbonUI XML (loaded via `CustomUI.bas`):
- XML must be embedded in the Excel file structure (requires Office Custom UI Editor or manual editing)
- Changes to Ribbon require Excel restart to take effect
- Maintain `myRibbon` object pointer for dynamic updates

## Common Patterns

### Date Handling
The system uses robust date detection:
- Handles Date values, text dates, and dates with time components
- Always uses `Int(dateValue)` to strip time portions for comparisons
- Header rows contain dates; `FindeDatumsspalte()` searches for matching columns

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

## Version Conventions

- Major versions (v2.x, v3.x): Significant feature additions
- Branch naming: `claude/description-vX.X-sessionID`
- Commits: German or English, descriptive of changes
- Release notes: Maintained in separate `RELEASE_NOTES_vX.X.md` file

## Important Notes

- **Excel-Specific**: This code only runs within Excel VBA environment
- **German Language**: All UI text, comments, and variable names are in German
- **No Standalone Tests**: Testing requires manual Excel workbook testing
- **Binary Files**: `.frx` files are binary and not human-readable
- **Worksheet CodeNames**: References like `Tabelle3` are CodeNames (not sheet names)
- **Manual Calculation**: Always consider recalculation implications when modifying UDFs

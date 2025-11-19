# Personalplaner - AI Assistant Guide

## Project Overview

**Personalplaner** (Personnel Planner) is an Excel VBA-based personnel management and scheduling system designed for Swiss organizations. The application manages employee schedules, absences, projects, workload calculations, and generates weekly reports (Wochenrapporte).

**Language**: German (Switzerland)
**Platform**: Microsoft Excel with VBA (Visual Basic for Applications)
**Version**: VBA7 compatible (64-bit Excel)

## Project Purpose

The system provides:
- Employee scheduling across calendar weeks (Kalenderwochen/KW)
- Absence tracking (vacations, sick days, military service, etc.)
- Project assignment and time tracking
- Workload and capacity calculations
- Automated weekly report generation
- Email notifications for report submissions
- Custom Excel ribbon UI for navigation

## Architecture Overview

### Core Structure

The project is organized in a **version-controlled VBA format** where each module, class, and form is stored as individual files:

```
Personalplaner/
├── DieseArbeitsmappe.doccls    # Workbook events and initialization
├── CustomUI.bas                 # Custom Excel Ribbon UI
├── Modul1-5.bas                # Legacy/utility modules
├── m*.bas                       # Feature modules (calendar, calculations, filtering)
├── UF_*.frm/frx                # UserForms and their binary data
├── Tabelle*.doccls             # Worksheet class modules
└── shWRTemplate.doccls         # Weekly report template sheet
```

### Key Technologies

- **VBA 7** (64-bit compatible with LongPtr declarations)
- **Scripting.Dictionary** for data collections and caching
- **ListObjects** (Excel Tables) for structured data
- **Outlook Integration** for email automation
- **Custom Ribbon XML** for UI customization
- **Named Ranges** for dynamic references

## File Organization and Responsibilities

### Workbook Events (`DieseArbeitsmappe.doccls`)

**Purpose**: Application initialization and global event handlers

**Key Features**:
- Sets calculation to Manual on workbook open for performance
- Shows/hides UserForms (UF_Filter, UF_Projekte) based on active sheet
- Refreshes custom ribbon on sheet activation

### Core Modules

#### `mKalender.bas` - Calendar Management
**@Folder**: "Personalplaner"

**Key Functions**:
- `ErstelleKalenderMitArbeitstagen()`: Creates calendar with weekdays only (Mon-Fri)
- `FerienUndFeiertageEintragen()`: Marks holidays and school vacations
- `BedingteFormatierungMitDropdownsInTabellen()`: Applies conditional formatting and dropdowns

**Important Constants**:
- `anzahlZeilen = 50`: Number of employee rows in weekly plans

**Absence Codes**:
- `F` = Ferien (Vacation)
- `Fx` = Ferien nicht bewilligt (Vacation not approved)
- `U` = Unfall (Accident)
- `K` = Krank (Sick)
- `WK` = Militär (Military service)
- `S` = Schule (School)
- `ÜK` = Überbetr. Kurs (Inter-company course)
- `T` = Teilzeit (Part-time)

#### `mBerechnung.bas` - Calculations and Formulas
**@Folder**: "FORMELN"

**Key UDFs (User Defined Functions)**:
- `VerweisMABAuslastungTotal()`: Retrieves workload values by date with column offset
- `FindeDatumsspalte()`: Robust date column finder (handles text and date formats)
- `AbwesendeMAB()`: Counts absent employees on a given date
- `ZaehleCodes()`: Counts cells with specific absence codes
- `AuslastungMitAusschluss()`: Calculates workload excluding specified absence types
- `VerfuegbareMitarbeiter()`: Counts available employees

**Important Notes**:
- Functions handle merged date headers
- Date comparisons ignore time portions using `Int(CDbl(date))`
- Supports both numeric dates and text date formats (dd.mm.yyyy)
- Uses `Application.Volatile True` for real-time calculations

#### `mAuslastung.bas` - Capacity Calculations
**@Folder**: "FORMELN"

**Purpose**: Specialized workload calculations referencing `Tabelle3` (main personnel sheet)

**Key Functions**:
- Similar to `mBerechnung.bas` but hardcoded to specific worksheets
- References `Tabelle3` by CodeName (not sheet name)

#### `mKWBlatt.bas` - Weekly Sheet Management
**@Folder**: "Personalplaner"

**Key Functions**:
- `NeuesKWBlattErstellen()`: Creates new weekly planning sheet from template
- `InitListBox()`: Populates ActiveX ListBox controls with unique values
- `AnfangsspalteVorherigeKW()`: Finds column of previous calendar week

**Workflow**:
1. Copies template sheet (`Tabelle7`)
2. Names sheet as `KW{number} {year}` (e.g., "KW15 2025")
3. Populates employee data from main planner
4. Applies conditional formatting
5. Initializes filter ListBoxes

#### `mWertesammler.bas` - Data Collection Utilities
**@Folder**: "Personalplaner"

**Key Functions**:
- `SammleEindeutigeWerteSchnell()`: Collects unique values from table columns
- `SammleEindeutigeWerteSchnellRng()`: Collects unique values from a range
  - Supports `includeHidden` parameter to skip hidden rows
  - Supports `OnlyFirstLine` to extract first line before line break (`Chr(10)`)

**Performance Optimizations**:
- Reads entire tables into arrays for speed
- Uses Dictionary for O(1) uniqueness checks
- Disables screen updating and events during processing

#### `mFormatierung.bas` - Formatting Utilities
**@Folder**: "FORMAT"

**Key Functions**:
- `ErsteZeileImBereichFett()`: Makes first line of multi-line cells bold

**Use Case**: Employee names with contact info on separate lines

#### `Modul5.bas` - Weekly Report Automation

**Key Procedures**:
- `WR_Anfordern()`: Sends email reminder to all employees for weekly report submission
  - Extracts email addresses from 3rd line of employee name cells (split by `vbNewLine`)
  - Creates single email with all recipients in TO field
  - Skips employees marked in column K (skip flag)

- `WR_Erstellen()`: Generates individual weekly reports for all employees
  - Creates new workbook with sheets per employee
  - Copies project data and hours from weekly plan
  - Handles special absences (Krank, Unfall, Militär, Ferien)
  - Prompts for project commission numbers and remarks
  - Saves as `Wochenrapporte_{KW}.xlsm`

**Important Data Structure**:
- Employee names stored as multi-line cells:
  ```
  Line 1: Display Name
  Line 2: Phone Number
  Line 3: Email Address
  ```

#### `mFilter.bas`, `mDatenüberprüfung.bas`, `Modul1-4.bas`

**Note**: These modules contain additional filtering, validation, and helper functions. Implementation details should be examined when working with filtering and data validation features.

### UserForms

#### `UF_Filter.frm` - Filtering Interface
**Purpose**: Provides filtering controls for the main planner view

#### `UF_Projekte.frm` - Project Selection
**Purpose**: Project picker for assigning employees to projects

**Key Features**:
- Loads projects from "Projektnummern" worksheet
- Double-click to insert project into active cell
- Validates cell is in correct column range (≥15 for Personalplaner, ≥5 for KW sheets)

#### `UF_ProjektErstellen.frm` - Project Creation
**Purpose**: Form for creating new projects

### Worksheet Classes

#### `Tabelle3.doccls` - Main Personnel Planner
**Purpose**: Primary yearly planning sheet named "Personalplaner"

**Structure**:
- Row 10: Date headers
- Columns A-N: Employee data (Number, Name, Function, Team, etc.)
- Columns O+: Daily assignments (one column per weekday)

#### `Tabelle7.doccls` - Weekly Template
**Purpose**: Template for KW (calendar week) sheets

#### `Tabelle8.doccls` - Employee Analysis
**Purpose**: "Auswertung Mitarbeiter" (Employee evaluation)

#### `shWRTemplate.doccls` - Weekly Report Template
**Purpose**: Template for individual employee weekly reports

### Custom UI (`CustomUI.bas`)

**@Folder**: "CustomUI"

**Purpose**: Manages Excel Ribbon customization

**Key Callbacks**:
- `OnLoad_PERSPLA()`: Initializes ribbon on workbook load
- `getVisible_PERSPLA()`: Controls ribbon tab visibility based on active sheet
- `onAction_PERSPLA()`: Handles button clicks

**Ribbon Buttons**:
- `TODAY`: Navigate to today's date column
- `ÜBERSICHT`: Show home view (Personalplaner)
- `AUSWERTUNG`: Show employee analysis dashboard
- `DIAGRAMM`: Show charts
- `FILTER`: Display filter UserForm
- `PROJEKT`: Display project selection
- `BERECHNEN`: Trigger manual calculation
- `PROJEKTEINGABE`: Show project input form
- `WP_SENDEN`: Send filtered PDFs to all employees
- `WR_ANFORDERUNG`: Send weekly report reminder emails
- `WR_ERSTELLEN`: Generate weekly reports

## Data Model

### Key Worksheets

1. **Personalplaner** (Tabelle3)
   - Main yearly planning grid
   - Employees in rows, dates in columns
   - Stores project assignments and absences

2. **KW{number} {year}** (Weekly sheets)
   - Created from Tabelle7 template
   - 5-day work week view (Mon-Fri)
   - Extracted data for specific calendar week

3. **Projektnummern** (wsProjekte)
   - Project master list
   - Columns: Project Name, Commission Number, Remarks

4. **Feiertage** (Holidays table in Tabelle1)
   - Holiday Name, Date

5. **Ferien** (School vacations table in Tabelle1)
   - Vacation Name, Start Date, End Date

6. **Auswertung Mitarbeiter** (Tabelle8)
   - Employee workload analysis and statistics

### Named Ranges

- `TAGE`: Range of date headers in calendar
- `RibbonID`: Pointer to IRibbonUI object for ribbon refresh

## Coding Conventions

### Naming Standards

**Variables**:
- Hungarian notation for objects: `ws` (Worksheet), `lo` (ListObject), `rng` (Range), `dict` (Dictionary)
- camelCase for local variables: `lastRow`, `dateCol`, `tempWert`
- PascalCase for parameters: `ByVal Datum As Date`

**Constants**:
- ALL_CAPS rare; usually PascalCase: `anzahlZeilen`

**Functions/Subs**:
- PascalCase: `ErstelleKalenderMitArbeitstagen`, `FindeDatumsspalte`
- German descriptive names

### Error Handling Pattern

```vba
Public Function Example() As Variant
    On Error GoTo ErrHandler

    ' ... code ...

    Example = result
    Exit Function

ErrHandler:
    Example = CVErr(xlErrValue)  ' Or appropriate error
End Function
```

**For Procedures with Cleanup**:
```vba
Public Sub Example()
    On Error GoTo ErrorHandler

    ' Store original settings
    Dim originalScreenUpdating As Boolean
    originalScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False

    ' ... code ...

CleanupAndExit:
    Application.ScreenUpdating = originalScreenUpdating
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = originalScreenUpdating
    MsgBox "Fehler " & Err.Number & ": " & Err.Description
    Resume CleanupAndExit
End Sub
```

### Performance Patterns

**Always used in heavy operations**:
```vba
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayAlerts = False

' ... operations ...

' Restore settings
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic ' Or keep Manual
Application.EnableEvents = True
Application.DisplayAlerts = True
```

**Array Processing**:
```vba
' Read entire range into array
Dim arr As Variant
arr = range.Value

' Process array (fast)
For i = 1 To UBound(arr, 1)
    For j = 1 To UBound(arr, 2)
        ' Process arr(i, j)
    Next j
Next i
```

### Multi-line Cell Handling

Employee data often uses Alt+Enter (`vbNewLine` / `Chr(10)`) for multi-line cells:

```vba
' Extract first line
Dim fullText As String
fullText = cell.Value
Dim firstLine As String
If InStr(fullText, Chr(10)) > 0 Then
    firstLine = Split(fullText, Chr(10))(0)
Else
    firstLine = fullText
End If
```

### CodeName vs. Name

**Always use CodeNames for reliability**:
```vba
Set ws = Tabelle3  ' CodeName - won't break if user renames sheet
' NOT: Set ws = Worksheets("Personalplaner")  ' User can rename
```

**CodeNames in this project**:
- `Tabelle3` = "Personalplaner"
- `Tabelle7` = KW template
- `Tabelle8` = "Auswertung Mitarbeiter"
- `Tabelle1` = Settings/reference data
- `shWRTemplate` = Weekly report template
- `wsProjekte` = Project list
- `Diagramm1` = Charts sheet

## Key Workflows

### Creating a New Calendar Week

1. User clicks a KW cell in Personalplaner sheet
2. `NeuesKWBlattErstellen()` is triggered
3. System checks if sheet "KW{n} {year}" already exists
4. If not, copies `Tabelle7` template
5. Populates with employee data for that week
6. Applies conditional formatting
7. Initializes filter ListBoxes

### Generating Weekly Reports

1. User navigates to a KW sheet
2. Clicks "WR_ERSTELLEN" ribbon button
3. `WR_Erstellen()` executes:
   - Collects unique projects from columns E:I
   - Prompts for commission numbers for new projects
   - Collects unique employees from column A
   - Creates new workbook `Wochenrapporte_{KW}.xlsm`
   - For each employee:
     - Copies `shWRTemplate`
     - Populates header (name, dates, KW)
     - Transfers daily project hours
     - Maps absences to special rows (26=Ferien, 27=Militär, 28=Unfall, 29=Krank)

### Sending Report Reminders

1. User clicks "WR_ANFORDERUNG" ribbon button
2. `WR_Anfordern()` executes:
   - Reads KW from active sheet
   - Collects employees from column A
   - Extracts email from 3rd line of name cells
   - Creates Outlook email with all recipients
   - Displays email for review (user must manually send)

## Development Guidelines for AI Assistants

### When Modifying Code

1. **Preserve Performance Optimizations**
   - Always maintain `ScreenUpdating = False` patterns
   - Keep array processing where used
   - Don't change `Calculation = xlCalculationManual` without discussion

2. **Error Handling**
   - Maintain existing error patterns
   - Always restore application settings in error handlers
   - Use `CleanupAndExit` pattern for resource cleanup

3. **Naming Conventions**
   - Continue using German function names (user expectation)
   - Keep Hungarian notation for object variables
   - Maintain `@Folder` annotations in module headers

4. **CodeNames**
   - Always use CodeNames (Tabelle3, etc.) not sheet names
   - Never hard-code sheet names in `Worksheets()` collections

5. **Multi-line Cell Data**
   - Remember cells can contain `Chr(10)` line breaks
   - Use `Split(text, Chr(10))` pattern consistently
   - Index: 0=Name, 1=Phone, 2=Email (for employee cells)

6. **Date Handling**
   - Always use `Int(CDbl(date))` to strip time portion for comparisons
   - Support both `Date` values and text dates (dd.mm.yyyy)
   - Use `WorksheetFunction.WeekNum(..., 2)` for ISO week numbers
   - Remember Monday = 1 in `Weekday(..., vbMonday)`

7. **ListObjects (Tables)**
   - Use `.DataBodyRange` for data (excludes headers)
   - Use `.HeaderRowRange` for headers
   - Use `.ListColumns("ColumnName")` for named access
   - Check `If Not lo.DataBodyRange Is Nothing` before processing

8. **Dictionary Usage**
   - Set `.CompareMode = vbTextCompare` for case-insensitive keys
   - Use for uniqueness checks and lookups
   - Return sorted dictionaries from collector functions

### Testing Considerations

1. **Hidden Rows**
   - Many functions skip hidden rows: `If Not row.Hidden Then`
   - Test with filtered data

2. **Empty/Error Cells**
   - Check `IsError(cell.Value2)` before processing
   - Use `Trim$(CStr(value))` pattern for safety

3. **Column References**
   - Personalplaner: Columns ≥15 (O+) are workdays
   - KW sheets: Columns ≥5 (E+) are workdays
   - This affects project assignment validation

4. **Calendar Assumptions**
   - Only weekdays (Mon-Fri) in calendars
   - Saturdays and Sundays are skipped
   - Row 10 always contains date headers

### Common Pitfalls

1. **Don't break merged cells** in calendar headers (KW, Month)
2. **Don't change column structure** without updating all dependent code
3. **Don't remove `Application.Volatile` from UDFs** without testing
4. **Don't use `ActiveSheet`** where specific sheet is needed - use CodeNames
5. **Don't forget to unlock UI** (`Application.ScreenUpdating = True`) on error

### Adding New Features

**Before implementing**:
1. Identify which module(s) should contain the code
2. Check if similar functionality exists elsewhere
3. Maintain the `@Folder` annotation style
4. Add descriptive `@Description` comments for public functions
5. Follow existing error handling patterns
6. Consider performance impact on large datasets (50+ employees, 260+ workdays)

**After implementing**:
1. Test with actual data (not just sample)
2. Verify it works with hidden rows/columns
3. Check it doesn't break when sheets are renamed
4. Ensure it handles empty cells and errors gracefully

## Git Workflow

This project uses VBA file extraction for version control:
- Each module is a separate `.bas` file
- Each UserForm is `.frm` (code) + `.frx` (binary)
- Each worksheet class is a `.doccls` file

**To export VBA**:
Use a VBA export tool or script to extract modules from the `.xlsm` file

**To import VBA**:
Import all modules back into a blank Excel workbook

## Dependencies

**Required References**:
- Microsoft Scripting Runtime (Scripting.Dictionary)
- Microsoft Outlook Object Library (for email automation)
- Microsoft Forms 2.0 Object Library (for UserForms)

**Excel Version**:
- Requires Excel 2010 or later (VBA7 with 64-bit support)
- Custom Ribbon requires Excel 2007+

## Performance Characteristics

**Typical Dataset**:
- 50-100 employees
- 260 workdays per year (52 weeks × 5 days)
- ~13,000 cells per table
- Multiple ListObjects per sheet

**Optimization Priorities**:
1. Array processing over cell-by-cell iteration
2. Dictionary caching over repeated searches
3. Minimal screen redraws
4. Manual calculation mode during bulk operations

## Localization

**Language**: Swiss German (de-CH)
- Uses Swiss date format: dd.mm.yyyy
- Weekday names in German: Montag, Dienstag, etc.
- Month names in German: Januar, Februar, etc.
- UI text and messages in German

**Encoding**: UTF-8 (supports umlauts: ä, ö, ü)

## Security Considerations

1. **Email addresses** are stored in cells (3rd line) - sensitive data
2. **Macros must be enabled** for application to function
3. **Outlook automation** requires user permission on first run
4. **File paths** are relative to workbook location

## Maintenance Notes

**Regular tasks**:
- Update "Feiertage" table annually with new holidays
- Update "Ferien" table with school vacation periods
- Review and clean old KW sheets periodically
- Backup project list ("Projektnummern") before major changes

**Known limitations**:
- Manual calculation mode requires explicit refresh
- Ribbon refresh requires Excel restart after major changes
- Multi-line cells can be accidentally broken by careless editing

---

## Quick Reference

### Finding Code

**Calendar operations**: `mKalender.bas`
**Date lookups**: `mBerechnung.bas`, `mAuslastung.bas`
**Weekly sheets**: `mKWBlatt.bas`
**Report generation**: `Modul5.bas`
**Filtering**: `mFilter.bas`, `UF_Filter.frm`
**Project selection**: `UF_Projekte.frm`
**Data collection**: `mWertesammler.bas`
**Formatting**: `mFormatierung.bas`
**Ribbon UI**: `CustomUI.bas`
**Workbook events**: `DieseArbeitsmappe.doccls`

### Key Sheet CodeNames

- `Tabelle3` - Main Personalplaner
- `Tabelle7` - KW template
- `Tabelle8` - Employee analysis
- `Tabelle1` - Settings/reference data
- `shWRTemplate` - Weekly report template
- `wsProjekte` - Project list
- `Diagramm1` - Charts

### Absence Code Reference

| Code | German | English |
|------|--------|---------|
| F | Ferien | Vacation |
| Fx | Ferien nicht bewilligt | Vacation not approved |
| U | Unfall | Accident |
| K | Krank | Sick |
| WK | Militär | Military service |
| S | Schule | School |
| ÜK | Überbetr. Kurs | Inter-company course |
| T | Teilzeit | Part-time |

---

**Last Updated**: 2025-11-19
**Version**: 1.0
**Maintained for**: AI Assistant Integration

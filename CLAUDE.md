# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

# Personalplaner - AI Assistant Guide

**Version**: 2.0 (Refactored)
**Last Updated**: 2025-11-19

## Quick Start

**What is this?** Excel VBA personnel management system for Swiss organizations with calendar scheduling, absence tracking, and automated reporting.

**Language:** German (Switzerland) | **Platform:** Excel VBA 7 (64-bit) | **Architecture:** Service-oriented with domain models

**Key Files:**
- `CalendarService.bas` - Calendar creation and management
- `WeeklySheetService.bas` - Weekly planning sheets (KW)
- `WeeklyReportService.bas` - Report generation and email reminders
- `Tabelle3.doccls` - Main Personalplaner worksheet
- `RibbonController.bas` - Custom Excel Ribbon UI
- `customUI.xml` - Ribbon configuration

**Critical Rules:**
1. Always use CodeNames (Tabelle3, not "Personalplaner")
2. Format weekday cells as TEXT first (`NumberFormat = "@"`)
3. Use `.HTMLBody` with UTF-8 for emails (not `.Body`)
4. Keep `Application.Calculation = xlCalculationManual` (never automatic)
5. Multi-line cells use `vbNewLine` (Chr(10)) - format: Name\nPhone\nEmail
6. Only ASCII characters in VBA code (no ä, ö, ü, emojis)
7. Calendar structure: Row 8=KW, Row 9=Month, Row 10=Weekdays (MO-FR)

**Testing:** All changes pushed to branch `claude/claude-md-mi5rpfamsyrm1wyu-01Wj9LJQ5P84ehTKbnzcyXLx`

---

## Project Overview

**Personalplaner** (Personnel Planner) is an Excel VBA-based personnel management and scheduling system designed for Swiss organizations. The application manages employee schedules, absences, projects, workload calculations, and generates weekly reports (Wochenrapporte).

**Language**: German (Switzerland)
**Platform**: Microsoft Excel with VBA (Visual Basic for Applications)
**Version**: VBA7 compatible (64-bit Excel)
**Architecture**: Refactored with clean service-oriented architecture and domain models

## Project Purpose

The system provides:
- **Yearly calendar creation** with work days only (Mon-Fri)
- **Employee scheduling** across calendar weeks (Kalenderwochen/KW)
- **Absence tracking** (vacations, sick days, military service, school, etc.)
- **Project assignment** and time tracking
- **Workload and capacity calculations** with Excel UDFs
- **Automated weekly report generation** (Wochenrapporte)
- **Email notifications** with UTF-8 encoding for German umlauts
- **Custom Excel Ribbon UI** for navigation and quick actions
- **Filtering and data analysis** with ActiveX ListBox controls

## Architecture Overview

### Core Structure (Refactored)

The project follows a **clean architecture** with separation of concerns:

```
Personalplaner/
├── Domain Models (Classes)
│   ├── AbsenceCode.cls          # Absence code domain model (F, K, U, WK, S, UK, T)
│   ├── Employee.cls             # Employee domain model with contact parsing
│   └── Project.cls              # Project domain model
│
├── Services (Business Logic)
│   ├── CalendarService.bas      # Calendar creation, holidays, formatting
│   ├── DateHelpers.bas          # Date calculations and dictionary sorting
│   ├── EmailService.bas         # PDF export and email sending
│   ├── EmployeeService.bas      # Employee data collection (optimized)
│   ├── FilterService.bas        # Table filtering with ListBox controls
│   ├── ProjectService.bas       # Project management and storage
│   ├── ValidationHelpers.bas    # Data validation utilities
│   ├── WeeklyReportService.bas  # Weekly report generation
│   ├── WeeklySheetService.bas   # KW sheet creation from template
│   └── WorkloadCalculations.bas # UDFs for Excel formulas
│
├── UI Layer
│   ├── RibbonController.bas     # Custom Ribbon callbacks
│   ├── UF_Filter.frm            # Filter form
│   ├── UF_Projekte.frm          # Project selection form
│   └── UF_ProjektErstellen.frm  # Project creation form
│
├── Worksheets
│   ├── DieseArbeitsmappe.doccls # Workbook events
│   ├── Tabelle3.doccls          # Main Personalplaner sheet
│   ├── Tabelle7.doccls          # KW template sheet
│   ├── Tabelle8.doccls          # Employee analysis sheet
│   ├── Diagramm1.doccls         # Chart sheet
│   └── shWRTemplate.doccls      # Weekly report template
│
└── Configuration
    └── customUI.xml             # Ribbon XML definition
```

### Key Technologies

- **VBA 7** (64-bit compatible with LongPtr declarations)
- **Scripting.Dictionary** for O(1) lookups and caching
- **ListObjects** (Excel Tables) for structured data management
- **Outlook Integration** (late binding) for email automation
- **Custom Ribbon XML** for UI customization
- **Named Ranges** ("TAGE", "RibbonID") for dynamic references
- **Array Processing** for performance (200+ employees over 5 years)
- **Rubberduck Annotations** (@Folder, @Description, @Param, @Returns)

## Domain Models

### AbsenceCode.cls

**@Folder**: "Domain.Models"

**Purpose**: Represents absence codes with colors and descriptions

**Properties**:
- `ShortForm`: "F", "K", "U", "WK", "S", "UK", "T"
- `LongForm`: Full text (e.g., "Ferien", "Krank", "Unfall")
- `Description`: Detailed explanation
- `ColorRGB`: RGB color for conditional formatting

**Key Method**:
- `GetAllCodes()`: Returns Dictionary of all absence codes

**Absence Codes**:
| Code | ShortForm | LongForm | Description | Color |
|------|-----------|----------|-------------|-------|
| F | F | Ferien | Vacation | Light Blue |
| Fx | Fx | Ferien nicht bewilligt | Vacation not approved | Orange |
| K | K | Krank | Sick | Yellow |
| U | U | Unfall | Accident | Red |
| WK | WK | Militär | Military service | Green |
| S | S | Schule | School | Light Green |
| UK | UK | Überbetr. Kurs | Inter-company course | Beige |
| T | T | Teilzeit | Part-time | Gray |

### Employee.cls

**@Folder**: "Domain.Models"

**Purpose**: Employee domain model with contact information parsing

**Properties**:
- `EmployeeNumber`: Unique employee ID
- `DisplayName`: Employee name (first line)
- `PhoneNumber`: Phone contact (second line)
- `EmailAddress`: Email address (third line)
- `Function`: Job function/role
- `Team`: Team assignment

**Key Methods**:
- `ParseFromCellValue(cellValue)`: Parses multi-line cell (Name\nPhone\nEmail)
- `ToMultiLineString()`: Converts back to multi-line format

**Note**: Employee cells use `vbNewLine` (Chr(10)) to separate lines

### Project.cls

**@Folder**: "Domain.Models"

**Purpose**: Project domain model with commission number and remarks

**Properties**:
- `ProjectName`: Project name/description
- `CommissionNumber`: Commission/order number
- `Remarks`: Additional notes

**Key Methods**:
- `Create()`: Factory method for creating instances
- `ToStorageString()`: Serializes to "CommissionNumber;Remarks" format
- `ParseFromStorageString()`: Deserializes from storage format

## Service Modules

### CalendarService.bas

**@Folder**: "Services.Calendar"
**Module**: Replaces old `mKalender.bas`

**Purpose**: Creates and manages calendar sheets with work days, holidays, and formatting

**Key Constants**:
- `EMPLOYEE_ROWS_COUNT = 50`: Number of employee rows
- `DATE_ROW_OFFSET = -1`: Date row relative to data start
- `CALENDAR_WEEK_ROW_OFFSET = -2`: KW row offset
- `MONTH_ROW_OFFSET = -3`: Month row offset

**Public Functions**:

#### `CreateWorkDayCalendar(startCell As Range)`
Creates a calendar with work days only (Monday-Friday)
- Prompts user for start/end date
- Clears existing calendar elements
- Creates columns for each weekday
- Formats cells with weekday names (MO, DI, MI, DO, FR) as TEXT
- Merges cells for KW, month, year headers
- Adds dotted borders between days, solid borders between weeks
- Extends all ListObjects to include calendar columns
- Sets column width to 2.0
- Applies conditional formatting and data validation dropdowns
- Optionally adds holidays and vacations

**IMPORTANT**: Date cells are formatted as TEXT (`NumberFormat = "@"`) to prevent Excel from interpreting "Di" as February date.

#### `AddHolidaysAndVacations()`
Adds holidays and school vacations to the calendar
- Reads from "Feiertage" and "Ferien" tables in Tabelle1
- Marks holidays with colored background
- Merges vacation period cells (properly unmerges first to avoid conflicts)

#### `ApplyConditionalFormattingToTables(Optional useShortForm, Optional startColumnIndex)`
Applies color-coded conditional formatting for absence codes
- Iterates through all ListObjects
- Creates formatting rules for each absence code
- Uses colors from AbsenceCode class

#### `ApplyDataValidationToTables(Optional startColumnIndex)`
Adds dropdown menus with absence codes
- Builds comma-separated list of all absence codes
- Applies Excel data validation to calendar columns
- Sets `InCellDropdown = True` for visible dropdown arrows

**Private Functions**:
- `ClearExistingCalendar()`: Deletes named range and clears area before recreation
- `ExtendListObjectsToCalendar()`: Extends all tables to last calendar column
- `FinalizeCalendarWeek()`: Merges and formats KW header
- `FinalizeMonth()`: Merges and formats month header
- `MarkVacationPeriod()`: Marks vacation period with merged cells
- `MarkHoliday()`: Marks single holiday
- `GetDateForColumn()`: Calculates actual date from weekday name and KW

### DateHelpers.bas

**@Folder**: "Services.Utilities"
**Module**: Replaces old `mBerechnung.bas` (partial) and `Modul2.bas`

**Purpose**: Date finding, formatting, and dictionary sorting utilities

**Public Functions**:

#### `FindDateColumn(targetSheet, headerRowNumber, searchDate, [firstColumnToSearch], [lastColumnToSearch]) As Long`
Robust date column finder (handles text and date formats)
- **Strategy 1**: Direct MATCH on numeric date serial
- **Strategy 2**: MATCH on text representation
- **Strategy 3**: Manual loop (handles dates with time, convertible text)
- Returns column number or 0 if not found

#### `FormatFirstLineBold(targetRange)`
Makes first line of multi-line cells bold
- Splits on `vbNewLine` (Chr(10))
- Sets bold formatting only for first line
- Used for employee names with contact info

#### `SortDictionaryAlphabetical(sourceDict) As Dictionary`
Sorts a Dictionary alphabetically by keys
- Returns new Dictionary with sorted keys
- Used for filter lists and data display

### EmployeeService.bas

**@Folder**: "Services.Employee"
**Module**: Replaces old `mWertesammler.bas`

**Purpose**: High-performance employee data collection optimized for 200+ employees over 5 years

**Public Functions**:

#### `GetUniqueValuesFromRange(targetRange, [includeHidden], [extractFirstLineOnly]) As Dictionary`
Collects unique values from a range
- **Performance**: Processes via array (fast) if includeHidden=True
- **Performance**: Uses row-by-row iteration if hidden rows excluded
- **extractFirstLineOnly**: Extracts only first line before `vbNewLine` (for employee names)
- Returns Dictionary with unique values as keys

#### `GetUniqueValuesFromListObjects([startColumn]) As Dictionary`
Collects unique values from all ListObjects on active sheet
- Merges data from all tables
- Optionally filters by column index
- Returns sorted Dictionary

**Important**: Uses Scripting.Dictionary for O(1) uniqueness checks

### ProjectService.bas

**@Folder**: "Services.Project"

**Purpose**: Project CRUD operations with user prompts

**Public Functions**:

#### `GetProjectSheet() As Worksheet`
Locates the "Projektnummern" worksheet
- Returns wsProjekte CodeName reference
- Used by WeeklyReportService to avoid "Blatt nicht gefunden" errors

#### `LoadProject(projectName) As Project`
Loads project from project sheet
- Searches for project by name
- Returns Project object or Nothing

#### `SaveProject(project As Project)`
Saves project to project sheet
- Adds new project or updates existing
- Stores in ListObject "Projektnummern"

#### `PromptForProjectDetails(projectName) As Project`
Prompts user for commission number and remarks
- Checks if project exists
- Asks user if they want to reuse existing data
- Returns Project object

### WeeklyReportService.bas

**@Folder**: "Services.WeeklyReport"
**Module**: Replaces old `Modul5.bas`

**Purpose**: Automated weekly report generation and email reminders

**Public Functions**:

#### `CreateWeeklyReports()`
Generates individual weekly reports for all employees
- Collects unique projects from columns E:I
- Prompts for commission numbers for new projects (uses ProjectService)
- Creates new workbook `Wochenrapporte_{KW}.xlsm`
- For each employee:
  - Copies shWRTemplate
  - Populates header (name, dates, KW)
  - Transfers daily project hours
  - Maps absences to special rows (26=Ferien, 27=Militär, 28=Unfall, 29=Krank)

#### `SendWeeklyReportReminder()`
Sends email reminder to all employees for weekly report submission
- Extracts email from 3rd line of employee name cells
- Creates Outlook email with all recipients in TO field
- **Uses HTMLBody with UTF-8** to properly display German umlauts
- Displays email for review (user must manually send)

**Important Email Encoding**:
```vba
'--- Correct UTF-8 encoding
emailBodyHTML = "<!DOCTYPE html>" & vbNewLine & _
                "<html>" & vbNewLine & _
                "<head>" & vbNewLine & _
                "<meta charset=""UTF-8"">" & vbNewLine & _
                "</head>" & vbNewLine & _
                "<body>..." & vbNewLine & _
                "</body></html>"

mailItem.HTMLBody = emailBodyHTML  '--- NOT .Body!
```

### WeeklySheetService.bas

**@Folder**: "Services.WeeklySheet"
**Module**: Replaces old `mKWBlatt.bas`

**Purpose**: Creates KW (calendar week) sheets from template

**Public Functions**:

#### `CreateWeeklySheet(selectedCell As Range)`
Creates new weekly planning sheet from template
- Parses calendar week number from cell
- Gets week start/end dates
- Checks if sheet already exists (if yes, activates it)
- Copies Tabelle7 template
- Names sheet as "KW{number} {year}"
- Populates employee data from main planner
- Replaces short absence codes with long form text
- Applies conditional formatting (useShortForm=False)
- Initializes filter ListBoxes
- **Refreshes Ribbon** to update TabWeeklyPlan visibility

#### `InitializeFilterListBox(targetSheet, listBoxName, columnName)`
Populates ActiveX ListBox with unique values from table column
- Used for "ListBoxFunktion" and "ListBoxTeam" filters
- Reads from first ListObject on sheet
- Extracts first line only (for employee names)

#### `FindPreviousWeekStartColumn(targetSheet) As Long`
Finds the starting column of the previous calendar week
- Used by filter to restrict visible data to recent dates
- Returns column number or 0 if today not found

**Private Functions**:
- `CopyEmployeeDataToWeeklySheet()`: Copies employee data for 5 weekdays
- `ReplaceAbsenceCodesWithLongForm()`: Replaces "F" with "Ferien", etc.

### WorkloadCalculations.bas

**@Folder**: "Services.Formulas"
**Module**: Replaces old `mBerechnung.bas`, `mAuslastung.bas`, `Modul3.bas`

**Purpose**: UDFs (User Defined Functions) for Excel formulas - workload, availability, day counting

**Public UDFs**:

#### `GetWorkloadByDate(targetDate, [columnOffset]) As Double`
Retrieves workload values by date with column offset
- Searches in Tabelle3 (main planner)
- Uses DateHelpers.FindDateColumn
- Supports column offset for accessing different data columns
- `Application.Volatile True` for real-time updates

#### `CalculateWorkload(targetDate, exclusionCriteria) As Double`
Calculates workload excluding specified absence types
- Parses exclusion criteria (semicolon-separated: "F;K;U")
- Counts employees excluding those with specified absence codes
- Used for capacity planning

#### `CountAvailableEmployees(targetDate, [columnOffset]) As Long`
Counts available employees on a given date
- Excludes all absence codes
- Used for resource allocation

#### `CountEmployeeDays(employeeName, filterCriteria) As Double`
Counts days matching criteria for an employee
- **filterCriteria = "Frei"**: Counts blank cells
- **filterCriteria = "Projekt"**: Counts cells with project codes (excludes absences)
- **filterCriteria = custom**: Semicolon-separated codes to count

**Performance Note**: All UDFs use `Application.Volatile True` and CodeName references (Tabelle3) for speed.

### EmailService.bas

**@Folder**: "Services.Email"
**Module**: New, extracted from `Modul4.bas`

**Purpose**: PDF export and email sending functionality

**Public Functions**:

#### `SendWeeklyPlanPDFToEmployees()`
Exports active sheet as PDF and creates email with attachments
- Exports to PDF in same directory as workbook
- Collects unique employee emails from filtered data
- Creates Outlook email with all recipients
- **Uses HTMLBody with UTF-8** for German umlauts
- Attaches PDF file
- Displays email for review

### FilterService.bas

**@Folder**: "Services.Filtering"
**Module**: Replaces old `mFilter.bas`

**Purpose**: Table filtering using ActiveX ListBox controls

**Public Functions**:

#### `ApplyTableFilter(targetSheet, listBoxName, columnName)`
Applies filter to first ListObject based on selected ListBox items
- Gets ActiveX ListBox from sheet
- Builds array of selected values
- Applies AutoFilter with `xlFilterValues` operator
- Used by "ListBoxFunktion" and "ListBoxTeam" in worksheet event handlers

### ValidationHelpers.bas

**@Folder**: "Services.Utilities"
**Module**: Replaces old `mDatenüberprüfung.bas`

**Purpose**: Data validation helper functions

**Public Functions**:
- `RemoveDataValidation(targetRange)`: Deletes validation from range
- `HasListValidation(targetCell) As Boolean`: Checks if cell has list validation
- `AddListValidation(targetRange, listItems)`: Adds dropdown validation

## UI Layer

### RibbonController.bas

**@Folder**: "UI.Ribbon"
**Module**: Replaces old `CustomUI.bas`

**Purpose**: Manages Custom Ribbon UI interactions and navigation

**CRITICAL**: Control IDs in this module must match customUI.xml ribbon configuration!

**Callback Functions** (EXACT signatures required):

#### `Sub OnLoad_PersonalPlaner(ribbon As IRibbonUI)`
Ribbon onLoad callback - initializes ribbon reference
- Stores ribbon pointer in named range "RibbonID"
- Sets `ribbonUI` module variable

#### `Sub OnRibbonButtonClick(control As IRibbonControl)`
Handles all button clicks via control.id
- **Navigation**: BtnGoToToday, BtnShowOverview, BtnShowDashboard, BtnShowChart
- **Filter/Projects**: BtnShowFilter, BtnShowProjects, BtnProjectInput
- **Settings**: BtnShowSettings (placeholder)
- **Calculation**: BtnRecalculate (uses ActiveSheet.Calculate for performance)
- **Weekly Reports**: BtnSendWeeklyPlan, BtnRequestWeeklyReports, BtnCreateWeeklyReports
- **Calendar**: BtnCreateCalendar, BtnOpenCurrentWeek (NEW!)

#### `Sub GetControlVisibility(control As IRibbonControl, ByRef returnedVal As Boolean)`
Controls visibility based on active sheet
- **TabDashboard**: Always visible (True)
- **TabWeeklyPlan**: Only visible for KW sheets (ActiveSheet.Name Like "KW*")

**Public Functions**:

#### `RefreshRibbon()`
Refreshes the ribbon UI
- Gets ribbon from pointer or direct reference
- Calls `ribbonUI.Invalidate`
- Used after creating KW sheets to update tab visibility

**Private Navigation Functions**:
- `NavigateToToday()`: Jumps to today's date in main planner
- `NavigateToOverview()`: Shows Personalplaner, hides others
- `NavigateToDashboard()`: Shows Auswertung Mitarbeiter
- `NavigateToChart()`: Shows Diagramm1
- `ShowProjectInput()`: Shows project creation form

**NEW Private Functions**:

#### `CreateNewCalendar()`
Creates new yearly calendar in Personalplaner
- Navigates to Personalplaner sheet
- Calls `Tabelle3.CreateYearlyCalendar`
- Shows success message

#### `OpenCurrentWeeklyPlan()`
Opens or creates the weekly plan for current calendar week
- Finds today's date in calendar (uses DateHelpers.FindDateColumn)
- Locates KW header cell in row 8
- Calls `WeeklySheetService.CreateWeeklySheet`
- **Reliable alternative to double-click** (works even if VBA errors occur)

### customUI.xml

**Purpose**: Excel Ribbon XML configuration

**Structure**:
```xml
<customUI xmlns="..." onLoad="OnLoad_PersonalPlaner">
  <ribbon>
    <tabs>
      <tab idQ="MARE:Mare-Tab" label="Maréchaux">

        <!-- Übersicht Group -->
        <group id="TabDashboard" getVisible="GetControlVisibility">
          <button id="BtnGoToToday" onAction="OnRibbonButtonClick" />
          <button id="BtnShowOverview" onAction="OnRibbonButtonClick" />
          <button id="BtnShowDashboard" onAction="OnRibbonButtonClick" />
          <button id="BtnShowChart" onAction="OnRibbonButtonClick" />
          <button id="BtnShowSettings" onAction="OnRibbonButtonClick" />
          <separator id="Separator1"/>
          <button id="BtnShowFilter" onAction="OnRibbonButtonClick" />
          <button id="BtnShowProjects" onAction="OnRibbonButtonClick" />
          <button id="BtnRecalculate" onAction="OnRibbonButtonClick" />
          <button id="BtnProjectInput" onAction="OnRibbonButtonClick" />
          <separator id="Separator2"/>
          <button id="BtnCreateCalendar" onAction="OnRibbonButtonClick" />
          <button id="BtnOpenCurrentWeek" onAction="OnRibbonButtonClick" />
        </group>

        <!-- Wochenplan Group -->
        <group id="TabWeeklyPlan" getVisible="GetControlVisibility">
          <button id="BtnSendWeeklyPlan" onAction="OnRibbonButtonClick" />
          <button id="BtnRequestWeeklyReports" onAction="OnRibbonButtonClick" />
          <button id="BtnCreateWeeklyReports" onAction="OnRibbonButtonClick" />
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>
```

**CRITICAL Callback Mapping**:
- **onLoad**: `OnLoad_PersonalPlaner` (NOT OnLoad_PERSPLA)
- **onAction**: `OnRibbonButtonClick` (NOT onAction_PERSPLA)
- **getVisible**: `GetControlVisibility` (NOT getVisible_PERSPLA)

### UserForms

#### UF_Filter.frm
**@Folder**: "UI.Forms"

**Purpose**: Filtering interface for main planner view

**Key Methods**:
- `LoadFilterData([startColumn])`: Loads unique values using EmployeeService
- Uses FilterService for applying filters

#### UF_Projekte.frm
**@Folder**: "UI.Forms"

**Purpose**: Project selection form

**Key Features**:
- `UserForm_Initialize()`: Auto-loads projects when opened (FIX for Ribbon issue)
- `LoadProjectData()`: Populates ListBox from project sheet
- Double-click to insert project into active cell
- Validates cell is in correct column range (≥15 for Personalplaner, ≥5 for KW sheets)

#### UF_ProjektErstellen.frm
**@Folder**: "UI.Forms"

**Purpose**: Form for creating new projects

## Worksheet Classes

### DieseArbeitsmappe.doccls

**@Folder**: "Core"

**Purpose**: Workbook event handlers

**Events**:

#### `Workbook_Open()`
- Sets `Application.Calculation = xlCalculationManual` (performance for 200+ employees)
- Enables events

#### `Workbook_SheetActivate()`
- Shows/hides UF_Filter and UF_Projekte based on active sheet
- Refreshes custom ribbon

### Tabelle3.doccls (Personalplaner)

**@Folder**: "Worksheets.MainPlanner"
**CodeName**: Tabelle3
**Sheet Name**: "Personalplaner"

**Purpose**: Main yearly planning sheet

**Constants**:
- `WEEKLY_INTERVAL = 5`: 5 workdays per week (Mon-Fri)
- `CALENDAR_START_CELL = "O10"`: First date cell in calendar header

**Structure**:
- Row 8: KW headers (merged across 5 columns)
- Row 9: Month headers (merged across variable columns)
- Row 10: Date/weekday names (MO, DI, MI, DO, FR)
- Columns A-N: Employee data (Number, Name, Function, Team, etc.)
- Columns O+: Daily assignments (one column per weekday)

**Public Methods**:

#### `CreateYearlyCalendar()`
Creates the yearly calendar with work days (Mon-Fri) starting at O10
- Calls `CalendarService.CreateWorkDayCalendar`

**Event Handlers**:

#### `Worksheet_Activate()`
Initializes filter and project forms
- Loads filter data starting from previous week
- Loads project data

#### `Worksheet_BeforeDoubleClick(Target, Cancel)`
Creates new KW sheet when user double-clicks
- **SCENARIO 1**: Direct KW header click (row 8 merged cell)
- **SCENARIO 2**: ListObject cell click (FIX: extended functionality!)
  - Finds column of clicked cell
  - Looks up corresponding KW header in row 8
  - Creates weekly sheet for that KW

#### `Worksheet_Change(Target)`
Handles recurring weekly entries for School (S) and Part-time (T)
- Prompts user if they want weekly recurrence
- Fills every 5th column (one work week)

**Private Methods**:
- `PromptAndApplyWeeklyRecurrence()`: Applies S/T to every week

### Tabelle7.doccls (KW Template)

**@Folder**: "Worksheets.WeeklyTemplate"
**CodeName**: Tabelle7

**Purpose**: Template for KW (calendar week) sheets

**Public Properties**:
- `copying As Boolean`: Flag to indicate sheet is being copied

**Event Handlers**:

#### `Worksheet_Activate()`
Initializes filter ListBoxes
- Calls `WeeklySheetService.InitializeFilterListBox` for "ListBoxFunktion" and "ListBoxTeam"

#### `ListBoxFunktion_Change()` and `ListBoxTeam_Change()`
Apply filters using FilterService
- Calls `FilterService.ApplyTableFilter`
- Triggers calculation

### Tabelle8.doccls (Auswertung Mitarbeiter)

**@Folder**: "Worksheets.EmployeeAnalysis"
**CodeName**: Tabelle8
**Sheet Name**: "Auswertung Mitarbeiter"

**Purpose**: Employee workload analysis and statistics

**Constants**:
- `MAIN_PLANNER_TABLE_NAME = "tblAL"`: Source table in Tabelle3
- `EVALUATION_TABLE_NAME = "tblAuswertung"`: Target table in Tabelle8
- `EMPLOYEE_NAME_COLUMN_OFFSET = 6`: Column 7 in table (0-based offset)

**Public Methods**:

#### `PopulateEmployeeEvaluation()`
Populates evaluation table with employee data from main planner
- Reads employee names from column 7 of tblAL
- Clears existing evaluation data
- Creates evaluation rows
- **Performance**: Reads entire column into array
- Initializes filter ListBoxes
- Recalculates formulas

**Event Handlers**:
- `ListBoxFunktion_Change()`: Filters by function
- `ListBoxMitarbeiter_Change()`: Filters by employee
- `ListBoxTeam_Change()`: Filters by team

### shWRTemplate.doccls (Wochenrapport Template)

**@Folder**: "PDFs"
**CodeName**: shWRTemplate

**Purpose**: Template for individual employee weekly reports

**Event Handlers**:

#### `Worksheet_Change(Target)`
Formats commission numbers as "XX XXX XXX"
- Triggers when user types 8 digits
- Inserts spaces: "12 345 678"

## Data Model

### Key Worksheets

1. **Personalplaner** (Tabelle3)
   - Main yearly planning grid
   - Employees in rows, dates in columns
   - Stores project assignments and absences
   - ListObjects contain employee data and calendar

2. **KW{number} {year}** (created from Tabelle7)
   - 5-day work week view (Mon-Fri)
   - Extracted data for specific calendar week
   - Long form absence codes instead of short codes

3. **Projektnummern** (wsProjekte)
   - Project master list
   - Columns: Project Name, Commission Number, Remarks
   - Used by ProjectService

4. **Feiertage** (in Tabelle1)
   - ListObject "Feiertage"
   - Columns: Holiday Name, Date

5. **Ferien** (in Tabelle1)
   - ListObject "Ferien"
   - Columns: Vacation Name, Start Date, End Date

6. **Auswertung Mitarbeiter** (Tabelle8)
   - Employee workload analysis and statistics
   - Uses formulas with WorkloadCalculations UDFs

### Named Ranges

- **TAGE**: Range of date cells in calendar (row 10)
- **RibbonID**: Pointer to IRibbonUI object (used by RefreshRibbon)

## Coding Conventions

### Naming Standards

**Variables**:
- Hungarian notation for objects: `ws` (Worksheet), `lo` (ListObject), `rng` (Range), `dict` (Dictionary)
- camelCase for local variables: `lastRow`, `dateColumn`, `employeeName`
- PascalCase for parameters: `ByVal targetDate As Date`

**Functions/Subs**:
- PascalCase: `CreateWorkDayCalendar`, `FindDateColumn`
- English names (changed from German in refactoring)

**Constants**:
- ALL_CAPS or PascalCase: `EMPLOYEE_ROWS_COUNT`, `WeeklyInterval`

### Rubberduck Annotations

All modules use Rubberduck annotations:

```vba
'@Folder("Services.Calendar")
'@ModuleDescription("Creates and manages calendar sheets")

'@Description("Creates calendar with work days only")
'@Param startCell The cell where calendar starts
'@Param endDate The last date to include
'@Returns True if successful
Public Function CreateCalendar(...) As Boolean
    '@Ignore EmptyStringLiteral
    '@Todo Implement year wrapping
End Function
```

**Available Annotations**:
- `@Folder("path")`: Organizes modules in Rubberduck
- `@ModuleDescription("text")`: Module purpose
- `@Description("text")`: Function/sub purpose
- `@Param name Description`: Parameter documentation
- `@Returns Description`: Return value documentation
- `@Ignore RuleId`: Suppress Rubberduck inspection
- `@Todo text`: Mark incomplete implementations

### Error Handling Pattern

**For Functions**:
```vba
Public Function Example() As Variant
    On Error GoTo ErrHandler

    '... code ...

    Example = result
    Exit Function

ErrHandler:
    Example = CVErr(xlErrValue)
End Function
```

**For Procedures with Cleanup**:
```vba
Public Sub Example()
    On Error GoTo ErrorHandler

    '--- Store original settings
    Dim originalScreenUpdating As Boolean
    originalScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False

    '... code ...

CleanupAndExit:
    Application.ScreenUpdating = originalScreenUpdating
    Exit Sub

ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical
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

'... operations ...

'--- Restore settings
Application.ScreenUpdating = True
Application.Calculation = xlCalculationManual  '--- Keep manual!
Application.EnableEvents = True
Application.DisplayAlerts = True
```

**Array Processing**:
```vba
'--- Read entire range into array (FAST)
Dim dataArray As Variant
dataArray = listObj.DataBodyRange.Value

'--- Process array in memory
For rowIndex = 1 To UBound(dataArray, 1)
    For colIndex = 1 To UBound(dataArray, 2)
        '--- Process dataArray(rowIndex, colIndex)
    Next colIndex
Next rowIndex
```

**Dictionary Caching**:
```vba
'--- O(1) uniqueness checks
Dim uniqueValues As Dictionary
Set uniqueValues = New Dictionary
uniqueValues.CompareMode = vbTextCompare  '--- Case-insensitive

For Each value In sourceRange
    If Not uniqueValues.Exists(value) Then
        uniqueValues.Add value, True
    End If
Next value
```

### Multi-line Cell Handling

Employee data uses Alt+Enter (`vbNewLine` / `Chr(10)`) for multi-line cells:

```vba
'--- Parsing multi-line cell
Dim lines() As String
lines = Split(cellValue, vbNewLine)
If UBound(lines) >= 0 Then displayName = lines(0)
If UBound(lines) >= 1 Then phoneNumber = lines(1)
If UBound(lines) >= 2 Then emailAddress = lines(2)

'--- Creating multi-line cell
cellValue = displayName & vbNewLine & phoneNumber & vbNewLine & emailAddress
```

### CodeName vs. Name

**Always use CodeNames for reliability**:
```vba
Set ws = Tabelle3  '--- CodeName (safe, won't break if user renames)
' NOT: Set ws = Worksheets("Personalplaner")  '--- User can rename!
```

**CodeNames in this project**:
| CodeName | Default Sheet Name |
|----------|-------------------|
| `Tabelle3` | "Personalplaner" |
| `Tabelle7` | KW template (hidden) |
| `Tabelle8` | "Auswertung Mitarbeiter" |
| `Tabelle1` | Settings/reference data |
| `shWRTemplate` | Weekly report template |
| `wsProjekte` | "Projektnummern" |
| `Diagramm1` | Charts sheet |

## Key Workflows

### Creating a New Calendar

1. User clicks **"Kalender"** ribbon button (or calls `Tabelle3.CreateYearlyCalendar`)
2. `RibbonController.CreateNewCalendar()` executes
3. Navigates to Personalplaner sheet
4. `CalendarService.CreateWorkDayCalendar()` is called
5. System prompts for start/end date (e.g., 01.01.2025 - 31.12.2025)
6. Clears existing calendar (deletes "TAGE" named range, clears area)
7. Creates columns for each weekday (Mon-Fri only)
8. Formats cells:
   - **Row 10**: Weekday names (MO, DI, MI, DO, FR) as TEXT
   - **Row 8**: KW numbers (merged across 5 columns)
   - **Row 9**: Month names (merged across variable columns)
9. Adds borders:
   - Dotted borders between individual days
   - Solid borders between weeks
10. Sets column width to 2.0
11. Extends all ListObjects to include calendar columns
12. Creates "TAGE" named range
13. Optionally adds holidays and vacations
14. Applies conditional formatting (color-coded absence codes)
15. Adds data validation dropdowns (absence codes)
16. Stays on Personalplaner sheet

### Creating a Weekly Sheet

**Method 1: Double-click** (original functionality + extended)
1. User double-clicks on KW header (row 8 merged cell) OR
2. User double-clicks on any ListObject data cell
3. `Tabelle3.Worksheet_BeforeDoubleClick` event fires
4. System finds KW header for that column
5. `WeeklySheetService.CreateWeeklySheet` is called

**Method 2: Ribbon button** (NEW! More reliable)
1. User clicks **"Akt. Woche"** ribbon button
2. `RibbonController.OpenCurrentWeeklyPlan()` executes
3. System finds today's date in calendar
4. System locates KW header for today's column
5. `WeeklySheetService.CreateWeeklySheet` is called

**Creation Process**:
1. Parses KW number and dates from header cell
2. Checks if sheet "KW{n} {year}" already exists
   - If yes: Activates existing sheet and exits
   - If no: Continues with creation
3. Copies Tabelle7 template
4. Names sheet as "KW{number} {year}" (e.g., "KW15 2025")
5. Populates header (KW, dates, timestamp)
6. Copies employee data for that week (5 columns)
7. Replaces short absence codes with long form text
8. Applies conditional formatting (useShortForm=False)
9. Initializes filter ListBoxes (Funktion, Team)
10. Refreshes Ribbon to show TabWeeklyPlan
11. Activates new sheet

### Generating Weekly Reports

1. User navigates to a KW sheet
2. User clicks **"WR erstellen"** ribbon button
3. `WeeklyReportService.CreateWeeklyReports()` executes:
   - Collects unique projects from columns E:I
   - Prompts for commission numbers for new projects (uses ProjectService)
   - Collects unique employees from column A
   - Creates new workbook `Wochenrapporte_{KW}.xlsm`
   - For each employee:
     - Copies `shWRTemplate`
     - Populates header (name, dates, KW)
     - Transfers daily project hours
     - Maps absences to special rows:
       - Row 26 = Ferien
       - Row 27 = Militär
       - Row 28 = Unfall
       - Row 29 = Krank

### Sending Report Reminders

1. User navigates to a KW sheet
2. User clicks **"WR einfordern"** ribbon button
3. `WeeklyReportService.SendWeeklyReportReminder()` executes:
   - Reads KW from active sheet
   - Collects employees from column A
   - Extracts email from 3rd line of name cells (splits by vbNewLine)
   - Creates Outlook email with all recipients in TO field
   - **Uses HTMLBody with UTF-8** to properly display German umlauts
   - Displays email for review (user must manually send)

## Development Guidelines for AI Assistants

### When Modifying Code

1. **Preserve Performance Optimizations**
   - Always maintain `ScreenUpdating = False` patterns
   - Keep array processing where used
   - Keep `Calculation = xlCalculationManual` (never change to Automatic)

2. **Error Handling**
   - Maintain existing error patterns
   - Always restore application settings in error handlers
   - Use `CleanupAndExit` pattern for resource cleanup

3. **Naming Conventions**
   - Use English function names (post-refactoring standard)
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
   - Support both `Date` values and text dates
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

9. **Character Encoding**
   - Only use ASCII characters in code comments
   - Replace: ä→ae, ö→oe, ü→ue in error messages
   - For emails: Use HTMLBody with UTF-8 charset, NOT .Body

10. **Text Formatting**
    - When displaying weekday names (Mo-Fr), format cells as TEXT first
    - Use `NumberFormat = "@"` before setting value
    - Prevents Excel from interpreting as dates

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
   - Row 8: KW headers (merged)
   - Row 9: Month headers (merged)
   - Row 10: Date/weekday names (TEXT format)

5. **Ribbon Callbacks**
   - EXACT signatures required (parameter names matter!)
   - `ribbon As IRibbonUI` (not `ByVal`)
   - `control As IRibbonControl` (not `ByVal`)
   - `ByRef returnedVal` (not `returnVisible`)

### Common Pitfalls

1. **Don't break merged cells** in calendar headers (KW, Month)
2. **Don't change column structure** without updating all dependent code
3. **Don't remove `Application.Volatile` from UDFs** without testing
4. **Don't use `ActiveSheet`** where specific sheet is needed - use CodeNames
5. **Don't forget to unlock UI** (`ScreenUpdating = True`) on error
6. **Don't use `.Body`** for emails - use `.HTMLBody` with UTF-8
7. **Don't format weekday names without setting NumberFormat="@" first**

### Adding New Features

**Before implementing**:
1. Identify which module(s) should contain the code
2. Check if similar functionality exists elsewhere
3. Maintain the `@Folder` annotation style
4. Add descriptive `@Description` comments for public functions
5. Follow existing error handling patterns
6. Consider performance impact on large datasets (200+ employees, 260+ workdays)

**After implementing**:
1. Test with actual data (not just sample)
2. Verify it works with hidden rows/columns
3. Check it doesn't break when sheets are renamed
4. Ensure it handles empty cells and errors gracefully
5. Update CLAUDE.md and REFACTORING.md

## Known Issues & Fixes

### Fixed Issues (as of 2025-11-19)

✅ **Conditional Formatting nicht angezeigt** - Fixed in CalendarService
✅ **Dropdown-Menü mit Absencecodes fehlt** - Added ApplyDataValidationToTables
✅ **Datumszellen als Datum interpretiert** - Now formatted as TEXT first
✅ **ListObject geht nicht bis Ende** - Simplified extension logic
✅ **Diagramm öffnet sich** - Changed Tabelle1.Activate to Tabelle3.Activate
✅ **Schulferien falsch gemerged** - Now unmerges before merging
✅ **Email-Kodierung falsch** - Uses HTMLBody with UTF-8 instead of Body
✅ **Projektnummern nicht gefunden** - Uses ProjectService.GetProjectSheet
✅ **UF_Projekte lädt nicht** - Added UserForm_Initialize
✅ **Doppelklick unzuverlässig** - Added Ribbon buttons as alternative

### Active TODOs

See REFACTORING.md for complete list of TODOs

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
- Weekday names: Mo, Di, Mi, Do, Fr (German abbreviations)
- Month names in German: Januar, Februar, etc.
- UI text and messages in German

**Encoding**: UTF-8 for emails, ASCII for VBA comments

## Security Considerations

1. **Email addresses** are stored in cells (3rd line) - sensitive data
2. **Macros must be enabled** for application to function
3. **Outlook automation** requires user permission on first run
4. **File paths** are relative to workbook location

## Quick Reference

### Finding Code

| What | Where |
|------|-------|
| Calendar operations | CalendarService.bas |
| Date lookups | DateHelpers.bas |
| Weekly sheets | WeeklySheetService.bas |
| Report generation | WeeklyReportService.bas |
| Email sending | EmailService.bas, WeeklyReportService.bas |
| Filtering | FilterService.bas |
| Project management | ProjectService.bas |
| Employee data collection | EmployeeService.bas |
| Workload UDFs | WorkloadCalculations.bas |
| Ribbon UI | RibbonController.bas |
| Workbook events | DieseArbeitsmappe.doccls |

### Key Sheet CodeNames

| CodeName | Sheet Name |
|----------|------------|
| Tabelle3 | Personalplaner |
| Tabelle7 | KW template (hidden) |
| Tabelle8 | Auswertung Mitarbeiter |
| Tabelle1 | Settings/reference data |
| shWRTemplate | Weekly report template |
| wsProjekte | Projektnummern |
| Diagramm1 | Charts |

### Absence Code Reference

| Code | German | English | Color |
|------|--------|---------|-------|
| F | Ferien | Vacation | Light Blue |
| Fx | Ferien nicht bewilligt | Vacation not approved | Orange |
| K | Krank | Sick | Yellow |
| U | Unfall | Accident | Red |
| WK | Militär | Military service | Green |
| S | Schule | School | Light Green |
| ÜK | Überbetr. Kurs | Inter-company course | Beige |
| T | Teilzeit | Part-time | Gray |

---

**Version**: 2.0 (Refactored)
**Last Updated**: 2025-11-19
**Maintained for**: AI Assistant Integration

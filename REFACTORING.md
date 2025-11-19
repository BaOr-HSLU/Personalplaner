# Personalplaner - Refactoring Dokumentation

## Übersicht

Die gesamte Codebase wurde refactoriert mit Fokus auf:
- ✅ **Rubberduck-Annotationen** (@Folder, @Description, @Todo)
- ✅ **Verständliche Variablen-/Parameternamen** (keine Abkürzungen wie MAB, lkey)
- ✅ **Performance-Optimierung** (200+ Mitarbeiter über 5 Jahre)
- ✅ **Modulare Architektur** (Klassen + Services)

## Neue Struktur

### Klassen (Domain Models)

| Datei | Beschreibung |
|-------|--------------|
| `AbsenceCode.cls` | Abwesenheitscodes (F, K, U, WK, etc.) mit Farben und Beschreibungen |
| `Employee.cls` | Mitarbeiter-Objekt mit Name, Email, Telefon, etc. |
| `Project.cls` | Projekt-Objekt mit Kommissionsnummer und Bemerkungen |

### Service-Module

| Datei | Alt | Beschreibung |
|-------|-----|--------------|
| `DateHelpers.bas` | mBerechnung.bas (teilweise) | Datums- und Kalender-Hilfsfunktionen |
| `EmployeeService.bas` | mWertesammler.bas | Mitarbeiter-Datensammlung (Performance-optimiert) |
| `ProjectService.bas` | - (neu) | Projektverwaltung und -speicherung |
| `CalendarService.bas` | mKalender.bas | Kalender-Erstellung und Formatierung |
| `WeeklyReportService.bas` | Modul5.bas | Wochenrapport-Erstellung und Email-Versand |
| `WeeklySheetService.bas` | mKWBlatt.bas | KW-Blatt-Erstellung aus Vorlage |
| `WorkloadCalculations.bas` | mBerechnung.bas, mAuslastung.bas | UDFs für Excel-Formeln (Auslastung, Verfügbarkeit) |

### UI-Module

| Datei | Alt | Beschreibung |
|-------|-----|--------------|
| `RibbonController.bas` | CustomUI.bas | Custom Ribbon Steuerung |
| `UF_Projekte.frm` | UF_Projekte.frm | Refactored mit besseren Namen |
| `DieseArbeitsmappe.doccls` | DieseArbeitsmappe.doccls | Workbook Events (aufgeräumt) |

## ⚠️ WICHTIG: CustomUI Ribbon Änderungen

Die Control-IDs im Ribbon XML müssen angepasst werden!

### Alte vs. Neue Control-IDs

| Alt | Neu | Typ |
|-----|-----|-----|
| `DASHBOARD` | `TabDashboard` | Tab |
| `WOCHENPLAN` | `TabWeeklyPlan` | Tab |
| `TODAY` | `BtnGoToToday` | Button |
| `ÜBERSICHT` | `BtnShowOverview` | Button |
| `AUSWERTUNG` | `BtnShowDashboard` | Button |
| `DIAGRAMM` | `BtnShowChart` | Button |
| `FILTER` | `BtnShowFilter` | Button |
| `PROJEKT` | `BtnShowProjects` | Button |
| `BERECHNEN` | `BtnRecalculate` | Button |
| `PROJEKTEINGABE` | `BtnProjectInput` | Button |
| `WP_SENDEN` | `BtnSendWeeklyPlan` | Button |
| `WR_ANFORDERUNG` | `BtnRequestWeeklyReports` | Button |
| `WR_ERSTELLEN` | `BtnCreateWeeklyReports` | Button |

### Callback-Änderungen

**Alt:**
```xml
<customUI onLoad="OnLoad_PERSPLA">
  <ribbon>
    <tabs>
      <tab id="DASHBOARD" getVisible="getVisible_PERSPLA">
        <group id="Navigation">
          <button id="TODAY" onAction="onAction_PERSPLA" />
```

**Neu:**
```xml
<customUI onLoad="OnLoad_PersonalPlaner">
  <ribbon>
    <tabs>
      <tab id="TabDashboard" getVisible="GetControlVisibility">
        <group id="Navigation">
          <button id="BtnGoToToday" onAction="OnRibbonButtonClick" />
```

## Namenskonventionen

### ✅ Gut - Neue Namen

```vba
Dim employeeName As String
Dim projectList As Dictionary
Dim weekStartDate As Date
Dim calendarWeekNumber As Long
```

### ❌ Schlecht - Alte Namen (entfernt)

```vba
Dim MABName As String
Dim PROJEKTE As Dictionary
Dim KWStart As Date
Dim KW As Long
```

## Performance-Optimierungen

Alle Performance-kritischen Bereiche wurden beibehalten:

### Array-Processing
```vba
'--- PERFORMANCE: Read entire table into array
tableData = listObj.DataBodyRange.Value
For rowIndex = 1 To UBound(tableData, 1)
    ' Process in memory (fast for 200+ employees)
Next
```

### Dictionary-Caching
```vba
'--- PERFORMANCE: Dictionary for O(1) lookups
Dim uniqueDict As Dictionary
Set uniqueDict = New Dictionary
```

### Application Settings
```vba
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
' ... operations ...
' Always restore afterwards!
```

## Rubberduck-Annotationen

Alle Module verwenden jetzt Rubberduck-Annotationen:

```vba
'@Folder("Services.Employee")
'@ModuleDescription("Manages employee data collection")

'@Description("Collects unique values from range")
'@Param targetRange The range to scan
'@Param includeHidden Whether to include hidden rows
Public Function GetUniqueValues(...)
    '@Ignore EmptyStringLiteral
    '@Todo Implement caching for large datasets
End Function
```

## Migration Guide

### Für Entwickler

1. **Importieren Sie alle neuen .bas und .cls Dateien** in Excel VBA
2. **Aktualisieren Sie die CustomUI XML-Datei** mit den neuen Control-IDs
3. **Testen Sie alle Ribbon-Buttons** nach der Aktualisierung
4. **Prüfen Sie Named Ranges**: `TAGE`, `RibbonID`

### Funktions-Mapping (Alte → Neue Namen)

| Alt | Neu |
|-----|-----|
| `FindeDatumsspalte()` | `DateHelpers.FindDateColumn()` |
| `SammleEindeutigeWerteSchnellRng()` | `EmployeeService.GetUniqueValuesFromRange()` |
| `ErsteZeileImBereichFett()` | `DateHelpers.FormatFirstLineBold()` |
| `WR_Erstellen()` | `WeeklyReportService.CreateWeeklyReports()` |
| `WR_Anfordern()` | `WeeklyReportService.SendWeeklyReportReminder()` |
| `ErstelleKalenderMitArbeitstagen()` | `CalendarService.CreateWorkDayCalendar()` |
| `NeuesKWBlattErstellen()` | `WeeklySheetService.CreateWeeklySheet()` |
| `VerweisMABAuslastungTotal()` | `WorkloadCalculations.GetWorkloadByDate()` |
| `AuslastungMitAusschluss()` | `WorkloadCalculations.CalculateWorkload()` |
| `VerfuegbareMitarbeiter()` | `WorkloadCalculations.CountAvailableEmployees()` |

## Testing Checklist

Nach dem Import testen:

- [ ] Ribbon-Buttons funktionieren alle
- [ ] Kalender erstellen funktioniert
- [ ] KW-Blatt erstellen funktioniert
- [ ] Wochenrapporte erstellen funktioniert
- [ ] Email-Erinnerungen senden funktioniert
- [ ] Projektauswahl-Form funktioniert
- [ ] UDFs in Formeln funktionieren (`=GetWorkloadByDate(...)`)
- [ ] Filter-Funktionen funktionieren
- [ ] Conditional Formatting wird korrekt angewendet

## Bekannte TODOs

```vba
'@Todo Implement SendFilteredPDFEmailToAll (in RibbonController.bas)
'@Todo Add caching for GetUniqueValuesFromRange (in EmployeeService.bas)
```

## Fragen oder Probleme?

Siehe `CLAUDE.md` für detaillierte Codebase-Dokumentation.

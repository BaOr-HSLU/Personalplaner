# Personalplaner - Refactoring Dokumentation

## Übersicht

Die gesamte Codebase wurde refactoriert mit Fokus auf:
- **Rubberduck-Annotationen** (@Folder, @Description, @Todo)
- **Verständliche Variabeln-/Parameternamen** (keine Abkürzungen wie MAB, lkey)
- **Performance-Optimierung** (200+ Mitarbeiter über 5 Jahre)
- **Modulare Architektur** (Klassen + Services)

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
| `DateHelpers.bas` | mBerechnung.bas (teilweise), Modul2.bas | Datums- und Kalender-Hilfsfunktionen, Dictionary-Sortierung |
| `EmployeeService.bas` | mWertesammler.bas | Mitarbeiter-Datensammlung (Performance-optimiert) |
| `ProjectService.bas` | - (neu) | Projektverwaltung und -speicherung |
| `CalendarService.bas` | mKalender.bas | Kalender-Erstellung und Formatierung |
| `WeeklyReportService.bas` | Modul5.bas | Wochenrapport-Erstellung und Email-Versand |
| `WeeklySheetService.bas` | mKWBlatt.bas | KW-Blatt-Erstellung aus Vorlage |
| `WorkloadCalculations.bas` | mBerechnung.bas, mAuslastung.bas, Modul3.bas | UDFs für Excel-Formeln (Auslastung, Verfügbarkeit, Stundenzähler) |
| `EmailService.bas` | Modul4.bas | PDF-Export und Email-Versand |
| `ValidationHelpers.bas` | mDatenüberprüfung.bas | Datenvalidierungs-Hilfsfunktionen |
| `FilterService.bas` | mFilter.bas | Table filtering mit ActiveX ListBox controls |

### UI-Module

| Datei | Alt | Beschreibung |
|-------|-----|--------------|
| `RibbonController.bas` | CustomUI.bas | Custom Ribbon Steuerung |
| `UF_Projekte.frm` | UF_Projekte.frm | Refactored mit besseren Namen |
| `UF_Filter.frm` | UF_Filter.frm, mFilter.bas | Refactored, verwendet neue Services |
| `DieseArbeitsmappe.doccls` | DieseArbeitsmappe.doccls | Workbook Events (aufgeräumt) |

## WICHTIG: CustomUI Ribbon Aenderungen

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

### Gut - Neue Namen

```vba
Dim employeeName As String
Dim projectList As Dictionary
Dim weekStartDate As Date
Dim calendarWeekNumber As Long
```

### Schlecht - Alte Namen (entfernt)

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
| `Stundenzähler()` | `WorkloadCalculations.CountEmployeeDays()` |
| `SendFilteredPDFEmailToAll()` | `EmailService.SendWeeklyPlanPDFToEmployees()` |
| `SortDictionaryAlphabetical()` | `DateHelpers.SortDictionaryAlphabetical()` |
| `EntferneDatenüberprüfung()` | `ValidationHelpers.RemoveDataValidation()` |
| `HasListValidation()` | `ValidationHelpers.HasListValidation()` |
| `ApplyTableFilter()` | `FilterService.ApplyTableFilter()` |
| `InitListBox()` | `WeeklySheetService.InitializeFilterListBox()` |

## Testing Checklist

Durchgefuehrt am 19.11.2025 13:22

- [x] Ribbon-Buttons funktionieren alle
  - **FEHLER**: Unbekannter Button: BtnShowSettings
  - **TODO**: Refresh soll nur das aktuelle Blatt neu rechnen um schnell zu bleiben

- [x] Kalender erstellen funktioniert
  - **FEHLER**: Wenn bereits ein Kalender besteht kommen Fehlermeldungen und die verbundenen Zellen werden komisch miteinander verbunden
  - **FEHLER**: Die Zellen mit dem Datum sind noch mit falschen Zeichenfolgen beschriftet, nicht mit dem Datum
  - **FEHLER**: Das ListObject welches mit den Mitarbeiter abgefuellt ist soll auf die neu erstellten Tage angepasst werden
  - **FEHLER**: Das Dropdown-Menue mit den Absencecodes fehlt. Bitte ergaenzen
  - **ZUSATZ**: Tage untereinander noch mit einer gestrichelten Linie unterteilen (nicht nur die Kalenderwochen)
  - **ZUSATZ**: Tage in der ausgewaehlten Zeile als MO, DI, MI, DO, FR formatieren ("TTT") und Spaltenbreite auf 2.0 erhoehen
  - **ZUSATZ**: Neuer Button um die ausgewaehlte Kalenderwoche als Wochenrapport zu oeffnen oder erstellen
  - **ZUSATZ**: Doppelklick-Verhalten aendern, so dass er auf allen Zellen im ListObject funktioniert

- [x] KW-Blatt erstellen funktioniert
  - **TODO**: Ribbon aktualisieren sobald das Blatt erstellt wurde

- [x] Wochenrapporte erstellen funktioniert
  - **FEHLER**: Blatt 'Projektnummern' nicht gefunden!
  - **INFO**: Die Rapporte werden trotz dem Fehler korrekt erstellt

- [x] Email-Erinnerungen senden funktioniert
  - **FEHLER**: Kodierungsproblem - ue wird als falsche Zeichen angezeigt

- [x] Projektauswahl-Form funktioniert
  - **FEHLER**: Es laedt die Projekte nicht, wenn ich es ueber den Ribbon oeffne

- [x] UDFs in Formeln funktionieren (`=GetWorkloadByDate(...)`)

- [x] Filter-Funktionen funktionieren

- [x] Conditional Formatting wird korrekt angewendet
  - **FEHLER**: Nein, sie wird gar nicht angezeigt

## Erkannte Probleme aus Tests

### Kritische Fehler (muessen behoben werden)

1. **CalendarService.bas** - Conditional Formatting wird nicht angezeigt
   - Problem: ApplyConditionalFormattingToTables wird nicht aufgerufen oder funktioniert nicht
   - Fix: Pruefen und korrigieren

2. **CalendarService.bas** - Dropdown-Menue mit Absencecodes fehlt
   - Problem: Data Validation wird nicht erstellt
   - Fix: AddDataValidationDropdown implementieren/korrigieren

3. **CalendarService.bas** - Datumszellen haben falsche Beschriftung
   - Problem: Format-String funktioniert nicht korrekt
   - Fix: Datums-Formatierung korrigieren

4. **CalendarService.bas** - ListObject wird nicht angepasst
   - Problem: Tabelle wird nicht auf neue Spalten erweitert
   - Fix: ListObject.Resize implementieren

5. **CalendarService.bas** - Fehlermeldungen bei bestehendem Kalender
   - Problem: Keine Pruefung ob Kalender bereits existiert
   - Fix: Bestehende Kalender-Elemente vor Neuerstellen loeschen

6. **WeeklyReportService.bas** - Blatt 'Projektnummern' nicht gefunden
   - Problem: wsProjekte CodeName wird nicht gefunden
   - Fix: CodeName pruefen und korrigieren

7. **Email-Kodierung** - Umlaute werden falsch dargestellt
   - Problem: Email-Body verwendet falsche Zeichenkodierung
   - Fix: Outlook HTMLBody mit UTF-8 verwenden

8. **UF_Projekte.frm** - Laedt Projekte nicht ueber Ribbon
   - Problem: LoadProjectData wird nicht aufgerufen oder hat Fehler
   - Fix: Event-Handler pruefen

9. **RibbonController.bas** - Unbekannter Button BtnShowSettings
   - Problem: Button existiert nicht oder Control-ID falsch
   - Fix: Button entfernen oder implementieren

### Verbesserungen / Zusatzfunktionen

10. **CalendarService.bas** - Gestrichelte Linien zwischen Tagen
    - Zusatz: Borders.LineStyle = xlDot zwischen einzelnen Tagen

11. **CalendarService.bas** - Tage als MO/DI/MI/DO/FR formatieren
    - Zusatz: Format(datum, "TTT") in Zeile 10

12. **CalendarService.bas** - Spaltenbreite 2.0
    - Zusatz: .ColumnWidth = 2.0 fuer Datumsspalten

13. **RibbonController.bas** - Neuer Button fuer Wochenrapport
    - Zusatz: BtnOpenWeeklyReport implementieren

14. **Tabelle3.doccls** - Doppelklick auf allen ListObject-Zellen
    - Zusatz: Pruefung ob Target in ListObject statt nur MergeCells

15. **WeeklySheetService.bas** - Ribbon nach KW-Blatt-Erstellung aktualisieren
    - Zusatz: RibbonController.RefreshRibbon aufrufen

16. **RibbonController.bas** - Refresh nur aktuelles Blatt
    - Zusatz: ActiveSheet.Calculate statt Application.Calculate

## Bekannte TODOs

```vba
'@Todo Fix: Conditional Formatting wird nicht angezeigt (CalendarService.bas)
'@Todo Fix: Dropdown-Menue mit Absencecodes fehlt (CalendarService.bas)
'@Todo Fix: Email-Kodierung fuer Umlaute (WeeklyReportService.bas, EmailService.bas)
'@Todo Fix: Projektnummern-Blatt nicht gefunden (WeeklyReportService.bas)
'@Todo Fix: UF_Projekte laedt nicht ueber Ribbon (RibbonController.bas)
'@Todo Feature: Gestrichelte Linien zwischen Tagen (CalendarService.bas)
'@Todo Feature: Tage als MO/DI/MI/DO/FR formatieren (CalendarService.bas)
'@Todo Feature: Button fuer Wochenrapport oeffnen (RibbonController.bas)
```

## Fragen oder Probleme?

Siehe `CLAUDE.md` für detaillierte Codebase-Dokumentation.

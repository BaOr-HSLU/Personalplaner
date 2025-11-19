# Entwicklungs-Dokumentation

Diese Dokumentation richtet sich an Entwickler, die am Personalplaner arbeiten möchten.

## 📋 Inhaltsverzeichnis

- [Architektur-Übersicht](#architektur-übersicht)
- [Module-Struktur](#module-struktur)
- [Entwicklungsumgebung](#entwicklungsumgebung)
- [Debugging](#debugging)
- [Häufige Aufgaben](#häufige-aufgaben)
- [Performance-Optimierung](#performance-optimierung)
- [Best Practices](#best-practices)

---

## Architektur-Übersicht

### High-Level Architektur

```
┌─────────────────────────────────────────────────┐
│                  Custom Ribbon UI               │
│              (CustomUI.bas)                     │
└──────────────────┬──────────────────────────────┘
                   │
        ┌──────────┴──────────┐
        │                     │
┌───────▼────────┐    ┌──────▼──────────┐
│  UserForms     │    │  Core Modules   │
│  - UF_Filter   │    │  - mKalender    │
│  - UF_Projekte │    │  - mBerechnung  │
│  - UF_Projekt  │    │  - mKWBlatt     │
│    Erstellen   │    │  - mAuslastung  │
└────────────────┘    └─────────┬───────┘
                                │
                      ┌─────────▼────────┐
                      │  Helper Modules  │
                      │  - mFormatierung │
                      │  - mWertesammler │
                      │  - mFilter       │
                      └──────────────────┘
```

### Datenfluss

```
User Input (Ribbon/UserForm)
    ↓
Event Handler (CustomUI.bas / UserForm Code)
    ↓
Core Function (mKalender, mBerechnung, etc.)
    ↓
Helper Functions (mFormatierung, mWertesammler)
    ↓
Excel Worksheets (Tabelle1-10, ListObjects)
    ↓
UI Update (Formatting, Refresh)
```

---

## Module-Struktur

### Core Modules

#### mKalender.bas
**Zweck:** Kalendererstellung und -verwaltung

**Hauptfunktionen:**
- `ErstelleKalenderMitArbeitstagen()` - Erstellt Kalender mit Mo-Fr
- `FerienUndFeiertageEintragen()` - Trägt Feiertage/Ferien ein
- `BedingteFormatierungMitDropdownsInTabellen()` - Formatierung

**Abhängigkeiten:**
- Tabelle1 (Stammdaten: Feiertage, Ferien)
- Named Range "TAGE"
- mFormatierung.bas

**Verwendung:**
```vb
' Kalender erstellen
Dim startZelle As Range
Set startZelle = Tabelle3.Range("O10")
ErstelleKalenderMitArbeitstagen startZelle
```

#### mBerechnung.bas
**Zweck:** Auslastungsberechnungen und UDFs

**Hauptfunktionen:**
- `VerweisMABAuslastungTotal()` - Datumbasierte Auslastung
- `FindeDatumsspalte()` - Robuste Datumssuche
- `AuslastungMitAusschluss()` - Auslastung mit Ausschlüssen
- `VerfuegbareMitarbeiter()` - Verfügbare Mitarbeiter zählen
- `ZaehleCodes()` - Abwesenheitscodes zählen

**Verwendung in Excel:**
```excel
=VerweisMABAuslastungTotal(A1; 0)
=AuslastungMitAusschluss(Ausschluss[Code])
=VerfuegbareMitarbeiter(Ausschluss[Code])
```

**Verwendung in VBA:**
```vb
Dim auslastung As Double
auslastung = VerweisMABAuslastungTotal(Date, 0)

Dim verfuegbar As Long
verfuegbar = VerfuegbareMitarbeiter(Range("A1:A10"))
```

#### mKWBlatt.bas
**Zweck:** Wochenplan-Erstellung

**Hauptfunktionen:**
- `NeuesKWBlattErstellen()` - Erstellt KW-Blatt
- `InitListBox()` - Befüllt ListBoxen
- `AnfangsspalteVorherigeKW()` - Findet vorherige KW

**Abhängigkeiten:**
- Tabelle7 (Vorlage)
- Tabelle3 (Personalplaner)
- mWertesammler.bas

#### CustomUI.bas
**Zweck:** Ribbon-Integration

**Callbacks:**
- `OnLoad_PERSPLA()` - Ribbon-Initialisierung
- `getVisible_PERSPLA()` - Sichtbarkeit von Tabs
- `onAction_PERSPLA()` - Button-Klicks

**Ribbon-Struktur:**
```xml
<customUI>
  <ribbon>
    <tabs>
      <tab id="DASHBOARD">
        <group id="grpNavigation">
          <button id="TODAY" />
          <button id="ÜBERSICHT" />
          <button id="AUSWERTUNG" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

### Helper Modules

#### mFormatierung.bas
- Formatierungshelfer
- Zellformatierung
- Bedingte Formatierung

#### mWertesammler.bas
- Datensammlung
- Dictionary-basierte Operationen
- Eindeutige Werte extrahieren

#### mFilter.bas
- Filterfunktionen
- Datenfilterung

---

## Entwicklungsumgebung

### Setup

1. **VBA-Editor öffnen**
   ```
   Alt + F11
   ```

2. **Project Explorer anzeigen**
   ```
   Strg + R
   ```

3. **Immediate Window anzeigen**
   ```
   Strg + G
   ```

4. **Properties Window**
   ```
   F4
   ```

### Nützliche Add-Ins

#### Rubberduck VBA (empfohlen)
- Code-Inspektionen
- Unit Testing
- Refactoring-Tools
- Code-Explorer

**Installation:**
https://rubberduckvba.com/

#### MZ-Tools
- Code-Templates
- Fehlerbehandlung-Generator
- Prozedur-Suche

### VBA-Referenzen

Folgende Referenzen sollten aktiviert sein:
- Visual Basic For Applications
- Microsoft Excel XX.0 Object Library
- Microsoft Office XX.0 Object Library
- Microsoft Scripting Runtime (für Dictionary)
- Microsoft Forms 2.0 Object Library (für UserForms)

**Aktivieren:**
VBA-Editor → Tools → References

---

## Debugging

### Debug.Print verwenden

```vb
Sub BeispielDebugging()
    Dim mitarbeiter As String
    mitarbeiter = "Max Mustermann"

    Debug.Print "Mitarbeiter:", mitarbeiter
    Debug.Print "Anzahl Zeichen:", Len(mitarbeiter)

    ' Im Immediate Window: Mitarbeiter: Max Mustermann
    '                      Anzahl Zeichen: 14
End Sub
```

### Breakpoints setzen

1. Klick auf grauen Rand neben Code-Zeile (roter Punkt erscheint)
2. Code ausführen
3. Pausiert bei Breakpoint
4. Mit F8 schrittweise durch Code gehen

### Locals Window

```
Ansicht → Lokal-Fenster
```

Zeigt alle lokalen Variablen und deren Werte.

### Watch Expressions

1. Rechtsklick auf Variable → "Überwachung hinzufügen"
2. Oder: Debug → Überwachung hinzufügen

### Error Handling debuggen

```vb
Sub DebugErrorHandling()
    On Error GoTo ErrHandler

    ' Fehler provozieren
    Dim x As Long
    x = 1 / 0

    Exit Sub

ErrHandler:
    Debug.Print "Fehler:", Err.Number, Err.Description
    Debug.Print "Quelle:", Err.Source
    Stop  ' Pausiert hier für Inspektion
End Sub
```

---

## Häufige Aufgaben

### Neues Modul hinzufügen

1. **Im VBA-Editor:**
   ```
   Einfügen → Modul
   ```

2. **Modul benennen:**
   ```
   F4 → Name: mMeinModul
   ```

3. **Header hinzufügen:**
   ```vb
   Attribute VB_Name = "mMeinModul"
   '@Folder "Personalplaner"
   '@ModuleDescription "Beschreibung des Moduls"
   Option Explicit
   ```

4. **In Git exportieren:**
   - Manuelle Export via VBA-Editor
   - Oder: VBA-Export-Tool verwenden

### Neue UserForm erstellen

1. **Einfügen → UserForm**

2. **Benennen:**
   ```
   F4 → Name: UF_MeinForm
   ```

3. **Controls hinzufügen:**
   - Toolbox verwenden
   - Properties setzen

4. **Event-Handler:**
   ```vb
   Private Sub btnOK_Click()
       ' Code hier
       Unload Me
   End Sub
   ```

### Ribbon-Button hinzufügen

1. **customUI.xml bearbeiten** (außerhalb VBA)

2. **Callback implementieren:**
   ```vb
   ' In CustomUI.bas
   Sub onAction_PERSPLA(control As IRibbonControl)
       Select Case control.ID
       Case "MEIN_BUTTON"
           MeineFunktion
       End Select
   End Sub
   ```

### Neue UDF erstellen

```vb
'@Description Beschreibung der Funktion
Public Function MeineUDF(ByVal input As Variant) As Variant
    On Error GoTo ErrHandler

    ' Implementierung
    MeineUDF = input * 2

    Exit Function

ErrHandler:
    MeineUDF = CVErr(xlErrValue)
End Function
```

**Verwendung in Excel:**
```excel
=MeineUDF(A1)
```

---

## Performance-Optimierung

### Standard-Optimierungen

```vb
Sub PerformanceOptimiert()
    ' Vorher: Status speichern
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEvents As Boolean

    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEvents = Application.EnableEvents

    ' Optimierungen aktivieren
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo Cleanup

    ' DEIN CODE HIER

Cleanup:
    ' Wiederherstellen
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEvents
End Sub
```

### Array-Operationen statt Zellen-Loops

**Langsam:**
```vb
For i = 1 To 10000
    Cells(i, 1).Value = i
Next i
```

**Schnell:**
```vb
Dim arr(1 To 10000, 1 To 1) As Variant
Dim i As Long
For i = 1 To 10000
    arr(i, 1) = i
Next i
Range("A1:A10000").Value = arr
```

### Dictionary statt Loops

**Langsam:**
```vb
For Each cell In rng
    If cell.Value = suchWert Then
        ' gefunden
        Exit For
    End If
Next
```

**Schnell:**
```vb
Dim dict As New Dictionary
' Vorher befüllen
For Each cell In rng
    dict(cell.Value) = cell.Row
Next

If dict.Exists(suchWert) Then
    ' gefunden
End If
```

### With-Statements verwenden

**Langsam:**
```vb
Range("A1").Value = "Test"
Range("A1").Font.Bold = True
Range("A1").Font.Size = 12
Range("A1").Interior.Color = RGB(255, 0, 0)
```

**Schnell:**
```vb
With Range("A1")
    .Value = "Test"
    With .Font
        .Bold = True
        .Size = 12
    End With
    .Interior.Color = RGB(255, 0, 0)
End With
```

---

## Best Practices

### 1. Immer Option Explicit

```vb
Option Explicit  ' Zwingt Variablen-Deklaration
```

### 2. Fehlerbehandlung

```vb
On Error GoTo ErrHandler
' Code
Exit Function/Sub
ErrHandler:
    ' Fehlerbehandlung
```

### 3. Ressourcen freigeben

```vb
Set objVariable = Nothing
```

### 4. Meaningful Names

```vb
' SCHLECHT
Dim x As Long

' GUT
Dim mitarbeiterAnzahl As Long
```

### 5. Kommentare

```vb
' Erkläre WARUM, nicht WAS
' SCHLECHT: "Schleife über Mitarbeiter"
' GUT: "Nur aktive Mitarbeiter berücksichtigen"
```

### 6. Funktionen klein halten

- Eine Funktion = Eine Aufgabe
- Max. 50-100 Zeilen
- Bei mehr: in kleinere Funktionen aufteilen

### 7. Magic Numbers vermeiden

```vb
' SCHLECHT
If status = 1 Then

' GUT
Const STATUS_AKTIV As Long = 1
If status = STATUS_AKTIV Then
```

---

## Nützliche Code-Snippets

### Safe String Conversion

```vb
Private Function SafeString(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        SafeString = vbNullString
    Else
        SafeString = CStr(v)
    End If
End Function
```

### Get Last Row

```vb
Function GetLastRow(ws As Worksheet, col As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function
```

### Progress Indicator

```vb
Sub ZeigeProgress(current As Long, total As Long)
    Dim percent As Long
    percent = Int((current / total) * 100)
    Application.StatusBar = "Fortschritt: " & percent & "%"
End Sub

' Am Ende:
Application.StatusBar = False
```

---

## Weitere Ressourcen

- [CONTRIBUTING.md](CONTRIBUTING.md) - Contribution Guidelines
- [README.md](README.md) - Projekt-Übersicht
- [RELEASE_NOTES_v2.7.md](RELEASE_NOTES_v2.7.md) - Detaillierte Features

---

**Happy Coding!** 🚀

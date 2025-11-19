# Contributing zu Personalplaner

Vielen Dank für dein Interesse an der Weiterentwicklung des Personalplaners! Diese Anleitung hilft dir, saubere und konsistente Beiträge zu leisten.

## 📋 Inhaltsverzeichnis

- [Code of Conduct](#code-of-conduct)
- [Erste Schritte](#erste-schritte)
- [Entwicklungs-Workflow](#entwicklungs-workflow)
- [Coding Standards](#coding-standards)
- [Commit-Konventionen](#commit-konventionen)
- [Pull Request Prozess](#pull-request-prozess)
- [VBA-spezifische Richtlinien](#vba-spezifische-richtlinien)
- [Testing](#testing)

---

## Code of Conduct

- Sei respektvoll und professionell
- Fokussiere auf konstruktives Feedback
- Hilf anderen bei Fragen und Problemen
- Dokumentiere deine Änderungen sauber

---

## Erste Schritte

### Voraussetzungen

- Microsoft Excel 2010 oder neuer (empfohlen: Excel 2016+)
- Git installiert und konfiguriert
- VBA-Editor-Kenntnisse
- Grundverständnis von Git/GitHub

### Repository klonen

```bash
git clone https://github.com/BaOr-HSLU/Personalplaner.git
cd Personalplaner
```

### Entwicklungsumgebung einrichten

1. Excel-Datei mit aktivierten Makros öffnen
2. VBA-Editor öffnen (`Alt + F11`)
3. Sicherstellen, dass "Trust access to the VBA project object model" aktiviert ist:
   - Datei → Optionen → Trust Center → Einstellungen für das Trust Center → Makroeinstellungen

---

## Entwicklungs-Workflow

### Branch-Strategie

Wir verwenden eine vereinfachte Git-Flow-Strategie:

```
main            ← Stable releases (v2.7, v2.8, etc.)
  ↑
  └── feature/* ← Feature-Entwicklung
  └── bugfix/*  ← Bugfixes
  └── hotfix/*  ← Kritische Hotfixes
```

### Neues Feature entwickeln

1. **Branch erstellen**
   ```bash
   git checkout main
   git pull origin main
   git checkout -b feature/dein-feature-name
   ```

2. **Entwickeln**
   - Code schreiben
   - Testen
   - Dokumentieren

3. **Committen**
   ```bash
   git add .
   git commit -m "feat: Beschreibung des Features"
   ```

4. **Push und PR**
   ```bash
   git push origin feature/dein-feature-name
   # Dann PR auf GitHub erstellen
   ```

### Bug fixen

1. **Branch erstellen**
   ```bash
   git checkout -b bugfix/beschreibung-des-bugs
   ```

2. **Fix implementieren und testen**

3. **Committen**
   ```bash
   git commit -m "fix: Behebe Problem mit XYZ"
   ```

---

## Coding Standards

### VBA-Namenskonventionen

```vb
' Module: mPascalCase (m-Präfix für Module)
' Beispiel: mKalender, mBerechnung

' Funktionen/Subs: PascalCase
Public Function BerechneAuslastung() As Double
Private Sub InitialisiereFormular()

' Variablen: camelCase
Dim mitarbeiterName As String
Dim anzahlTage As Long

' Konstanten: UPPERCASE_SNAKE_CASE
Const ANZAHL_ZEILEN As Long = 50
Const MAX_MITARBEITER As Long = 100

' UserForms: UF_PascalCase
' Beispiel: UF_Filter, UF_Projekte

' Controls: controlTypePascalCase
' Beispiel: btnSpeichern, txtMitarbeiterName, lstTeams
```

### Code-Kommentare

```vb
'==================================================================================================
'@Description Kurze Beschreibung der Funktion/Sub
'
'@Param paramName Type : Beschreibung des Parameters
'@Return Type: Beschreibung des Rückgabewerts
'==================================================================================================
Public Function MeineFunktion(ByVal paramName As String) As Long
    On Error GoTo ErrHandler

    ' Implementierung

    Exit Function

ErrHandler:
    MeineFunktion = 0
End Function
```

### Error Handling

**Immer** Error Handling implementieren:

```vb
Public Function BeispielFunktion() As Variant
    On Error GoTo ErrHandler

    ' Code hier

    Exit Function

ErrHandler:
    BeispielFunktion = CVErr(xlErrValue)
    ' Optional: Debug.Print oder Logging
End Function
```

### Performance Best Practices

```vb
Sub PerformanceOptimiert()
    ' Screen Updating ausschalten
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo Cleanup

    ' Dein Code hier

Cleanup:
    ' Immer wiederherstellen!
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

---

## Commit-Konventionen

Wir verwenden [Conventional Commits](https://www.conventionalcommits.org/):

### Format

```
<type>(<scope>): <subject>

<body>

<footer>
```

### Types

- **feat**: Neues Feature
- **fix**: Bugfix
- **docs**: Dokumentation
- **style**: Formatierung (keine Code-Änderung)
- **refactor**: Code-Refactoring
- **perf**: Performance-Verbesserung
- **test**: Tests hinzufügen/ändern
- **chore**: Build-Prozess, Dependencies, etc.

### Beispiele

```bash
# Feature
git commit -m "feat(kalender): Füge Feiertags-Import aus CSV hinzu"

# Bugfix
git commit -m "fix(auslastung): Korrigiere Berechnung bei Teilzeit-Mitarbeitern"

# Dokumentation
git commit -m "docs(readme): Aktualisiere Installationsanleitung"

# Refactoring
git commit -m "refactor(ribbon): Extrahiere Button-Handler in separate Methoden"

# Performance
git commit -m "perf(berechnung): Verwende Dictionary statt Array-Suche"
```

### Multi-Line Commits

```bash
git commit -m "feat(export): Füge Excel-Export für Auswertungen hinzu

- Neues Modul mExport.bas erstellt
- Export-Funktion für gefilterte Daten
- Automatische Formatierung des Exports
- Dateinamen-Generator mit Timestamp

Closes #42"
```

---

## Pull Request Prozess

### PR erstellen

1. **Beschreibender Titel**
   ```
   feat(kalender): Feiertags-Import aus CSV
   ```

2. **PR-Template ausfüllen**
   - Beschreibung der Änderungen
   - Verwandte Issues verlinken
   - Screenshots (bei UI-Änderungen)
   - Test-Plan

3. **Checklist abhaken**
   - [ ] Code folgt den Coding Standards
   - [ ] Alle Tests laufen durch
   - [ ] Dokumentation aktualisiert
   - [ ] CHANGELOG.md aktualisiert
   - [ ] Keine Debug.Print-Statements im finalen Code

### Review-Prozess

- Mindestens 1 Approval erforderlich
- Alle Kommentare müssen aufgelöst sein
- CI/CD-Checks müssen grün sein (falls vorhanden)

### Merge

- **Feature-Branches**: "Squash and merge" oder "Merge commit"
- **Hotfixes**: "Merge commit" mit Verweis auf Issue

---

## VBA-spezifische Richtlinien

### Module-Organisation

```
/Modules
  /Core
    - mKalender.bas      (Hauptfunktionalität)
    - mBerechnung.bas
  /UI
    - CustomUI.bas
    - UF_Filter.frm
  /Helpers
    - mFormatierung.bas
    - mWertesammler.bas
```

### Option Explicit

**Immer** `Option Explicit` verwenden:

```vb
'@Folder "Personalplaner"
Option Explicit
```

### Folder-Attribute

Rubberduck VBA-Annotations verwenden:

```vb
'@Folder "Personalplaner"
'@ModuleDescription "Kalendererstellung und Formatierung"
Option Explicit
```

### Vermeiden

```vb
' NICHT:
Dim x, y, z  ' Alle werden als Variant deklariert

' BESSER:
Dim x As Long
Dim y As Long
Dim z As String
```

### Verwenden

```vb
' GUT: Frühe Bindung
Dim dict As Scripting.Dictionary
Set dict = New Scripting.Dictionary

' VERMEIDEN: Späte Bindung (außer notwendig)
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
```

---

## Testing

### Manuelles Testing

Vor jedem Commit:

1. **Funktionstest**
   - Feature/Fix testen
   - Edge Cases prüfen

2. **Regressionstest**
   - Bestehende Features testen
   - Keine neuen Bugs eingeführt

3. **Performance-Test**
   - Bei großen Datenmengen testen
   - Keine merkbare Verlangsamung

### Test-Checklist für Releases

- [ ] Kalender erstellen (verschiedene Datumsbereiche)
- [ ] Abwesenheiten eintragen und formatieren
- [ ] Auslastungsberechnungen korrekt
- [ ] Wochenplan erstellen und exportieren
- [ ] Filter funktionieren
- [ ] Projektverwaltung funktioniert
- [ ] Ribbon-UI lädt korrekt
- [ ] Alle Buttons funktionieren
- [ ] Excel-Datei öffnet ohne Fehler

---

## Versionierung

### Semantic Versioning

```
MAJOR.MINOR.PATCH

2.7.0 → 2.8.0  (neues Feature)
2.7.0 → 2.7.1  (Bugfix)
2.7.0 → 3.0.0  (Breaking Change)
```

### CHANGELOG aktualisieren

Bei jeder Änderung `CHANGELOG.md` aktualisieren:

```markdown
## [Unreleased]

### Hinzugefügt
- Neues Feature X

### Behoben
- Bug Y in Modul Z
```

---

## Fragen?

Bei Fragen oder Unklarheiten:
- Issue erstellen mit Label `question`
- In der PR-Diskussion nachfragen
- Dokumentation prüfen (README.md, RELEASE_NOTES)

---

**Vielen Dank für deinen Beitrag zum Personalplaner!** 🚀

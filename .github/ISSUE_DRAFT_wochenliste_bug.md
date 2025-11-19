# Bug Report: Wochenliste kann nicht erstellt/geöffnet werden

---
**Issue Typ:** Bug Report
**Titel:** [BUG] Wochenliste kann nicht erstellt und geöffnet werden
**Labels:** bug, high-priority, wochenplan
**Betroffenes Modul:** mKWBlatt.bas

---

## Beschreibung des Problems

Die Wochenliste (KW-Blatt) kann nicht erstellt oder geöffnet werden. Beim Versuch, eine Kalenderwoche zu öffnen oder ein neues KW-Blatt zu erstellen, tritt ein Fehler auf oder es passiert nichts.

## Schritte zur Reproduktion

1. Öffne den Personalplaner
2. Navigiere zum Hauptblatt "Personalplaner"
3. Klicke auf eine Kalenderwoche (KW-Zelle)
4. Versuche, das KW-Blatt zu erstellen/öffnen
5. Fehler tritt auf

**Oder alternativ:**

1. Verwende Ribbon-Button für Wochenplan
2. Wähle Kalenderwoche aus
3. Klicke auf "Wochenplan erstellen"
4. Fehler oder keine Aktion

## Erwartetes Verhalten

- Ein neues Arbeitsblatt mit dem Namen "KW{Nummer} {Jahr}" sollte erstellt werden
- Das Blatt sollte basierend auf der Vorlage (Tabelle7) befüllt werden
- Mitarbeiterdaten sollten automatisch übertragen werden
- Das neue Blatt sollte aktiviert und sichtbar sein

## Tatsächliches Verhalten

Beschreibe genau, was stattdessen passiert:
- [ ] Keine Reaktion (nichts passiert)
- [ ] VBA-Fehlermeldung erscheint (welche?)
- [ ] Blatt wird erstellt, aber bleibt leer
- [ ] Excel stürzt ab
- [ ] Sonstiges: ___________

## Fehlermeldung

Falls eine VBA-Fehlermeldung erscheint, bitte hier einfügen:

```vb
[Fehlermeldung und Fehlercode hier einfügen, z.B.:
"Laufzeitfehler '1004':
Die Methode 'Copy' des Objekts '_Worksheet' ist fehlgeschlagen"]
```

**Fehlerhafte Zeile/Funktion:**
```vb
[Zeile aus mKWBlatt.bas, z.B.:
.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)]
```

## Screenshots

Falls möglich, füge Screenshots hinzu:
- Screenshot der Fehlermeldung
- Screenshot des Hauptblatts mit markierter KW
- Screenshot des VBA-Debuggers (falls im Debug-Modus)

## Umgebung

- **Excel-Version:** [z.B. Excel 2016, Excel 365, Excel 2019]
- **Betriebssystem:** [z.B. Windows 10, Windows 11]
- **Personalplaner-Version:** v2.7.0
- **Makros aktiviert:** [Ja/Nein]
- **VBA-Projekt-Zugriff:** [Aktiviert/Deaktiviert]

## Betroffenes Modul

- [x] Kalenderverwaltung
- [ ] Personalplanung
- [ ] Auslastungsberechnung
- [ ] Ribbon UI
- [x] Wochenplan-Export
- [ ] Filter
- [ ] Projektverwaltung
- [ ] Dashboard
- [ ] Sonstiges: ___________

## Betroffene Funktionen

**Hauptfunktion:** `NeuesKWBlattErstellen()` in mKWBlatt.bas (Zeile 10-131)

**Mögliche betroffene Bereiche:**
- Vorlage Tabelle7 (shWRTemplate) existiert nicht oder ist beschädigt
- Tabelle7.Visible ist auf xlSheetVeryHidden gesetzt
- Target-Range ist ungültig
- KW-Nummer kann nicht extrahiert werden
- Copy-Operation schlägt fehl
- ListObject in Vorlage fehlt

## Zusätzlicher Kontext

### Mögliche Ursachen

1. **Vorlage nicht vorhanden**
   - Tabelle7 (shWRTemplate) existiert nicht
   - Vorlage wurde versehentlich gelöscht

2. **Sichtbarkeits-Problem**
   - Tabelle7 ist auf xlSheetVeryHidden gesetzt
   - Kann nicht sichtbar gemacht werden

3. **Target-Range ungültig**
   - Ausgewählte Zelle enthält keine KW-Nummer
   - Datumsbereich für KW nicht gefunden

4. **ListObject fehlt**
   - Tabelle in Vorlage existiert nicht
   - Range("A7").ListObject schlägt fehl

5. **Berechtigungen**
   - VBA-Projekt-Zugriff nicht aktiviert
   - Makros nicht vollständig aktiviert

6. **Copying-Flag**
   - Tabelle7.copying Property verursacht Fehler
   - Rekursion durch Event-Handler

## Debugging-Schritte durchgeführt

- [ ] VBA-Debugger verwendet (F8 schrittweise)
- [ ] Immediate Window Ausgaben geprüft
- [ ] Locals Window inspiziert
- [ ] Breakpoint in NeuesKWBlattErstellen gesetzt
- [ ] Vorlage Tabelle7 auf Existenz geprüft
- [ ] Makro-Sicherheitseinstellungen geprüft

## Workaround

Falls ein temporärer Workaround existiert:

```vb
' Möglicher Workaround:
' 1. Tabelle7 manuell sichtbar machen
Tabelle7.Visible = xlSheetVisible

' 2. Funktion manuell im Immediate Window aufrufen
Call NeuesKWBlattErstellen(ActiveCell)
```

## Mögliche Lösung

Falls bereits eine Idee für die Lösung besteht:

### Option 1: Prüfung der Vorlage
```vb
' In NeuesKWBlattErstellen() zu Beginn hinzufügen:
If Tabelle7 Is Nothing Then
    MsgBox "Fehler: Vorlage nicht gefunden!", vbCritical
    Exit Sub
End If
```

### Option 2: Error Handling verbessern
```vb
' Besseres Error Handling:
On Error GoTo ErrHandler

' ... Code ...

ErrHandler:
    MsgBox "Fehler beim Erstellen des KW-Blatts: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
```

### Option 3: Sichtbarkeit prüfen
```vb
' Vorlage sichtbar machen falls nötig:
If Tabelle7.Visible <> xlSheetVisible Then
    Tabelle7.Visible = xlSheetVisible
End If
```

## Priorität

- [ ] Niedrig (nice-to-have)
- [ ] Mittel (würde Arbeit erleichtern)
- [x] Hoch (wichtig für täglichen Gebrauch)
- [ ] Kritisch (blockiert aktuelle Arbeit)

**Begründung:** Wochenplan-Export ist eine Kernfunktion des Personalplaners. Wenn diese nicht funktioniert, ist ein wichtiges Feature nicht nutzbar.

## Reproduzierbarkeit

- [ ] Tritt immer auf (100%)
- [ ] Tritt häufig auf (>50%)
- [ ] Tritt manchmal auf (<50%)
- [ ] Tritt selten auf
- [ ] Konnte nicht reproduziert werden

## Verwandte Issues

- Keine bekannt

## Checklist für Fix

- [ ] Problem identifiziert
- [ ] Lösung implementiert
- [ ] Unit-Tests hinzugefügt (falls möglich)
- [ ] Manuell getestet
- [ ] Code-Review durchgeführt
- [ ] CHANGELOG.md aktualisiert
- [ ] Commit mit "Fixes #XX"

---

**Bitte fülle die fehlenden Informationen aus, insbesondere:**
- Genaue Fehlermeldung (falls vorhanden)
- Excel-Version
- Screenshots
- Welches Verhalten genau auftritt
- Reproduzierbarkeit

**Für schnellere Diagnose hilfreich:**
- VBA-Debugger verwenden (F8)
- Im Immediate Window testen: `? Tabelle7.Name`
- Prüfen: Existiert ein Blatt namens "WRTemplate" oder ähnlich?

# Personalplaner Release v2.7

**Veröffentlichungsdatum:** 19. November 2025
**Projekttyp:** Excel VBA Personalplanungssystem
**Lizenz:** Proprietär

---

## Zusammenfassung

Release v2.7 stellt eine vollständige und stabile Version des Personalplaners dar. Das System bietet umfassende Funktionalität zur Verwaltung von Mitarbeiterressourcen, Abwesenheiten und Auslastungsplanung mit intuitiver Benutzeroberfläche.

---

## Hauptfunktionen

### 📅 **Kalenderverwaltung**
- Automatische Erstellung von Arbeitstageskalendern (Montag-Freitag)
- Übersichtliche Darstellung nach Kalenderwochen, Monaten und Jahren
- Integration von Feiertagen und Schulferien
- Flexible Datumsbereiche für Jahresplanung
- Visuell strukturierte Kalenderformatierung mit Rahmen und Merged Cells

### 👥 **Personalplanung**
- Verwaltung von Mitarbeiterdaten (Name, Funktion, Team, Kontaktdaten)
- Abwesenheitsverwaltung mit standardisierten Codes:
  - **F** = Ferien
  - **Fx** = Ferien nicht bewilligt
  - **K** = Krank
  - **U** = Unfall
  - **WK** = Militär
  - **S** = Schule
  - **ÜK** = Überbetrieblicher Kurs
  - **T** = Teilzeit
- Farbcodierte bedingte Formatierung für alle Abwesenheitstypen
- Filterbare Mitarbeiteransichten nach Team und Funktion

### 📊 **Auslastungsberechnung**
- **Robuste UDFs (User Defined Functions)** für Auslastungsberechnungen
- Automatische Berechnung verfügbarer Mitarbeiter unter Berücksichtigung von Abwesenheiten
- `VerweisMABAuslastungTotal()`: Datumbasierte Auslastungsabfrage mit Offset-Funktionalität
- `AuslastungMitAusschluss()`: Berechnung der Auslastungsquote mit konfigurierbaren Ausschlusskriterien
- `VerfuegbareMitarbeiter()`: Zählt verfügbare Mitarbeiter für einen bestimmten Tag
- `AbwesendeMAB()`: Zählt abwesende Mitarbeiter
- `ZaehleCodes()`: Flexibles Zählen von Abwesenheitscodes
- Unterstützung für verschiedene Datumsformate (Datumswerte, Text, mit Zeitanteil)

### 📑 **Wochenplan-Funktionalität**
- Automatische Erstellung von KW-spezifischen Arbeitsblättern
- Export von Wochenplänen basierend auf Vorlagen
- `NeuesKWBlattErstellen()`: Kopiert und befüllt KW-Blätter mit aktuellen Mitarbeiterdaten
- PDF-Export und E-Mail-Versand von gefilterten Wochenplänen
- Langform-Anzeige von Abwesenheitscodes in Wochenplänen

### 🎨 **Custom Ribbon UI**
- Intuitive Bedienung über benutzerdefinierte Excel-Menüleiste
- Schnellzugriff auf wichtigste Funktionen:
  - **Heute**: Springt zum aktuellen Datum
  - **Übersicht**: Hauptansicht
  - **Auswertung**: Dashboard mit Statistiken
  - **Diagramm**: Visualisierungen
  - **Filter**: Filterung nach Kriterien
  - **Projekt**: Projektverwaltung
  - **Berechnen**: Manuelle Neuberechnung
- Kontextsensitive Ribbon-Elemente (unterschiedliche Ansicht für KW-Blätter)
- Persistentes Ribbon über `myRibbon` Object-Pointer

### 🔍 **Filter & Projektverwaltung**
- UserForm-basierte Filterdialoge (`UF_Filter`)
- Projektverwaltung mit dediziertem UserForm (`UF_Projekte`)
- Projekterstellung mit Formular (`UF_ProjektErstellen`)
- Dynamische ListBox-Befüllung für Teams und Funktionen
- Eindeutige Werte-Sammlung mit Dictionary-basiertem Ansatz

### 📈 **Auswertungen & Dashboard**
- Dediziertes Auswertungsblatt für Mitarbeiterstatistiken
- Diagrammblatt für visuelle Darstellungen
- Automatische Berechnung bei Auswertungsaktivierung
- Statusleisten-Feedback für Benutzeraktionen

### ⚙️ **Leistungsoptimierung**
- Manuelle Berechnungseinstellung für bessere Performance
- `Application.ScreenUpdating = False` während intensiver Operationen
- Event-Handler-Management zur Vermeidung von Rekursion
- Effiziente Dictionary-basierte Lookup-Operationen

---

## Technische Details

### Architektur
- **Plattform:** Microsoft Excel (VBA7 und Legacy VBA kompatibel)
- **Sprache:** Visual Basic for Applications (VBA)
- **Module:** 15 VBA-Module (.bas, .doccls, .frm)
- **ListObjects:** Tabellenbasierte Datenverwaltung mit strukturierten Referenzen

### Code-Module
| Modul | Beschreibung |
|-------|--------------|
| `mKalender.bas` | Kalendererstellung und Formatierung |
| `mBerechnung.bas` | Auslastungsberechnungen und UDFs |
| `mAuslastung.bas` | Zusätzliche Auslastungsfunktionen |
| `mKWBlatt.bas` | Wochenplan-Erstellung |
| `mFilter.bas` | Filterfunktionalität |
| `mFormatierung.bas` | Formatierungsroutinen |
| `mWertesammler.bas` | Datensammlung und -aggregation |
| `CustomUI.bas` | Ribbon-Integration |
| `DieseArbeitsmappe.doccls` | Workbook-Event-Handler |
| `UF_Filter.frm` | Filter-UserForm |
| `UF_Projekte.frm` | Projekt-UserForm |
| `UF_ProjektErstellen.frm` | Projekterstellungs-UserForm |

### Wichtige Funktionen
```vba
' Hauptfunktionen
Sub ErstelleKalenderMitArbeitstagen(ByVal startZelle As Range)
Public Function VerweisMABAuslastungTotal(ByVal Datum As Date, Optional ByVal offset As Long = 0) As Double
Public Function AuslastungMitAusschluss(ByVal rngAusschluss As Range, Optional ByVal abteilung = False) As Double
Public Function VerfuegbareMitarbeiter(ByVal rngAusschluss As Range, Optional ByVal abteilung = False) As Long
Public Function FindeDatumsspalte(ByVal ws As Worksheet, ByVal HeaderRow As Long, ByVal Suchdatum As Date) As Long
Public Sub NeuesKWBlattErstellen(Target As Range)
Sub BedingteFormatierungMitDropdownsInTabellen(Optional ByVal Kurzform As Boolean = True)
Public Sub FerienUndFeiertageEintragen()
```

---

## Installation & Verwendung

### Systemanforderungen
- Microsoft Excel 2010 oder neuer (empfohlen: Excel 2016+)
- Makros müssen aktiviert sein
- VBA7 oder kompatible Version

### Erste Schritte
1. Excel-Datei mit aktivierten Makros öffnen (.xlsm)
2. Beim ersten Öffnen wird die Berechnung auf "Manuell" gestellt (Performance-Optimierung)
3. Custom Ribbon wird automatisch geladen
4. Navigation über Ribbon-Buttons oder Blatt-Aktivierung

### Kalender erstellen
1. Gewünschte Startzelle auswählen
2. Makro `ErstelleKalenderMitArbeitstagen` ausführen
3. Start- und Enddatum eingeben
4. Optional: Feiertage automatisch eintragen lassen

### Wochenplan erstellen
1. Kalenderwoche im Hauptblatt auswählen (Zelle mit KW-Nummer)
2. Makro `NeuesKWBlattErstellen` aufrufen
3. Automatische Befüllung mit Mitarbeiterdaten
4. Export als PDF möglich

---

## Bekannte Einschränkungen

- Kalender berücksichtigt nur Montag-Freitag (Werktage)
- Feiertage müssen in der Tabelle "Feiertage" gepflegt sein
- Schulferien müssen in der Tabelle "Ferien" gepflegt sein
- Maximale Mitarbeiteranzahl durch Excel-Zeilenlimit beschränkt (50 Zeilen konfiguriert)
- Ribbon wird erst nach Excel-Neustart vollständig aktualisiert bei Änderungen

---

## Wartung & Support

### Datenbanktabellen
Folgende Tabellen müssen gepflegt werden:
- **Feiertage**: Name, Datum
- **Ferien**: Name, Start-Datum, End-Datum
- **Mitarbeiter**: Nummer, Name, Funktion, Team, Kontaktdaten

### Performance-Tipps
- Berechnung bleibt auf "Manuell" für große Datenmengen
- Bei Bedarf über Ribbon "Berechnen" oder `F9` neu berechnen
- `Application.ScreenUpdating` wird automatisch gesteuert

---

## Changelog

### Version 2.7 (2025-11-19)
**Umfassende Erstveröffentlichung mit vollständigem Funktionsumfang**

#### Neu implementiert:
- Vollständige Kalenderverwaltung mit Arbeitstagen
- Robuste Datumserkennung mit Unterstützung verschiedener Formate
- Umfassende Auslastungsberechnungen mit UDFs
- Custom Ribbon UI mit kontextsensitiven Elementen
- Wochenplan-Export mit automatischer Befüllung
- Feiertags- und Ferienintegration
- Bedingte Formatierung mit Farbcodierung
- Filter- und Projektverwaltungs-UserForms
- PDF-Export und E-Mail-Versand
- Dashboard mit Auswertungen
- Performance-Optimierungen

#### Technische Verbesserungen:
- Dictionary-basierte Lookup-Operationen für bessere Performance
- Error-Handling in allen kritischen Funktionen
- VBA7 und Legacy VBA Kompatibilität
- Event-Handler-Management zur Vermeidung von Rekursionen
- Statusleisten-Feedback für Benutzeraktionen

---

## Mitwirkende

Entwickelt für die effiziente Personalplanung und Ressourcenverwaltung.

---

## Lizenz

Proprietär - Alle Rechte vorbehalten

---

**Ende der Release Notes v2.7**

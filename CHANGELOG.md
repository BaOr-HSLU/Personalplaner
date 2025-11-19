# Changelog

Alle nennenswerten Änderungen an diesem Projekt werden in dieser Datei dokumentiert.

Das Format basiert auf [Keep a Changelog](https://keepachangelog.com/de/1.0.0/),
und dieses Projekt folgt [Semantic Versioning](https://semver.org/lang/de/).

## [Unreleased]

### Geplant
- TBD

---

## [2.7.0] - 2025-11-19

### Hinzugefügt

#### Kalenderverwaltung
- Automatische Erstellung von Arbeitstageskalendern (Montag-Freitag)
- Übersichtliche Darstellung nach Kalenderwochen, Monaten und Jahren
- Integration von Feiertagen aus Tabelle "Feiertage"
- Integration von Schulferien aus Tabelle "Ferien"
- Flexible Datumsbereiche für Jahresplanung
- Visuell strukturierte Kalenderformatierung mit Rahmen und Merged Cells
- Named Range "TAGE" für Datumsbereich

#### Personalplanung
- Verwaltung von Mitarbeiterdaten (Nummer, Name, Funktion, Team, Kontaktdaten)
- Abwesenheitsverwaltung mit standardisierten Codes:
  - F = Ferien
  - Fx = Ferien nicht bewilligt
  - K = Krank
  - U = Unfall
  - WK = Militär
  - S = Schule
  - ÜK = Überbetrieblicher Kurs
  - T = Teilzeit
- Farbcodierte bedingte Formatierung für alle Abwesenheitstypen
- Filterbare Mitarbeiteransichten nach Team und Funktion

#### Auslastungsberechnung
- `VerweisMABAuslastungTotal()` - Datumbasierte Auslastungsabfrage mit Offset
- `FindeDatumsspalte()` - Robuste Datumserkennung (numerisch, Text, mit Zeitanteil)
- `AuslastungMitAusschluss()` - Auslastungsquote mit konfigurierbaren Ausschlüssen
- `VerfuegbareMitarbeiter()` - Zählt verfügbare Mitarbeiter für Tag
- `AbwesendeMAB()` - Zählt abwesende Mitarbeiter
- `ZaehleCodes()` - Flexibles Zählen von Abwesenheitscodes
- Unterstützung für verschiedene Datumsformate

#### Custom Ribbon UI
- Benutzerdefinierte Excel-Menüleiste mit IRibbonUI
- Schnellzugriff-Buttons:
  - Heute - Springt zum aktuellen Datum
  - Übersicht - Hauptansicht
  - Auswertung - Dashboard
  - Diagramm - Visualisierungen
  - Filter - Filterung aktivieren
  - Projekt - Projektverwaltung
  - Berechnen - Manuelle Neuberechnung
- Kontextsensitive Ribbon-Elemente (KW-Blätter vs. Hauptansicht)
- Persistentes Ribbon über myRibbon Object-Pointer
- Automatische Ribbon-Aktualisierung bei Blatt-Wechsel

#### Wochenplan-Funktionalität
- `NeuesKWBlattErstellen()` - Automatische KW-Blatt-Erstellung
- Export von Wochenplänen basierend auf Vorlagen (Tabelle7)
- Automatische Befüllung mit aktuellen Mitarbeiterdaten
- PDF-Export-Funktionalität (`SendFilteredPDFEmailToAll()`)
- E-Mail-Versand von Wochenplänen
- Langform-Anzeige von Abwesenheitscodes in KW-Blättern
- Dynamische ListBox-Befüllung für Teams und Funktionen

#### Filter & Projektverwaltung
- UserForm `UF_Filter` - Filterdialog
- UserForm `UF_Projekte` - Projektverwaltung
- UserForm `UF_ProjektErstellen` - Projekterstellung
- `InitListBox()` - Automatische ListBox-Befüllung
- Dictionary-basierte eindeutige Werte-Sammlung

#### Dashboard & Auswertungen
- Dediziertes Auswertungsblatt (Tabelle8)
- Diagrammblatt (Diagramm1) für Visualisierungen
- Automatische Berechnung bei Auswertungsaktivierung
- Statusleisten-Feedback für Benutzeraktionen

#### Dokumentation
- README.md mit Projekt-Übersicht und Schnellstart
- RELEASE_NOTES_v2.7.md mit vollständiger Feature-Dokumentation
- Inline-Code-Dokumentation mit @Description-Attributen
- Referenztabelle für Abwesenheitscodes

#### Technische Verbesserungen
- VBA7 und Legacy VBA Kompatibilität (PtrSafe Declarations)
- Dictionary-basierte Lookup-Operationen für Performance
- Umfassendes Error-Handling in allen kritischen Funktionen
- Event-Handler-Management zur Vermeidung von Rekursionen
- Manuelle Berechnungseinstellung für bessere Performance
- `Application.ScreenUpdating = False` während intensiver Operationen
- Bedingte Formatierung mit `FormatConditions.Add`
- Merge-Cell-Management für Kalenderdarstellung

### Geändert
- Berechnung auf "Manuell" gestellt (Performance)
- Tabelle7 (Vorlage) wird dynamisch ein-/ausgeblendet

### Module-Struktur
- **mKalender.bas** - Kalendererstellung und Formatierung
- **mBerechnung.bas** - Auslastungsberechnungen und UDFs
- **mAuslastung.bas** - Zusätzliche Auslastungsfunktionen
- **mKWBlatt.bas** - Wochenplan-Erstellung
- **mFilter.bas** - Filterfunktionalität
- **mFormatierung.bas** - Formatierungsroutinen
- **mWertesammler.bas** - Datensammlung und -aggregation
- **CustomUI.bas** - Ribbon-Integration
- **DieseArbeitsmappe.doccls** - Workbook-Event-Handler
- **UF_Filter.frm** - Filter-UserForm
- **UF_Projekte.frm** - Projekt-UserForm
- **UF_ProjektErstellen.frm** - Projekterstellungs-UserForm

---

## Versioning-Schema

Dieses Projekt verwendet [Semantic Versioning](https://semver.org/lang/de/):

- **MAJOR** (X.0.0): Inkompatible API-Änderungen oder grundlegende Umstrukturierung
- **MINOR** (0.X.0): Neue Funktionalität, abwärtskompatibel
- **PATCH** (0.0.X): Bugfixes, abwärtskompatibel

### Beispiele:
- `2.7.0` → `2.8.0`: Neue Features (z.B. neues Dashboard)
- `2.7.0` → `2.7.1`: Bugfix (z.B. Fehler in Datumsberechnung)
- `2.7.0` → `3.0.0`: Breaking Change (z.B. Umstellung auf andere Datenstruktur)

---

## Changelog-Kategorien

### Hinzugefügt (Added)
Für neue Features.

### Geändert (Changed)
Für Änderungen an bestehender Funktionalität.

### Veraltet (Deprecated)
Für Features, die bald entfernt werden.

### Entfernt (Removed)
Für entfernte Features.

### Behoben (Fixed)
Für Bugfixes.

### Sicherheit (Security)
Für Sicherheits-relevante Änderungen.

---

[Unreleased]: https://github.com/BaOr-HSLU/Personalplaner/compare/v2.7.0...HEAD
[2.7.0]: https://github.com/BaOr-HSLU/Personalplaner/releases/tag/v2.7.0

# Personalplaner v2.7

Ein umfassendes Excel-VBA basiertes Personalplanungssystem für effiziente Ressourcenverwaltung und Auslastungsplanung.

![Version](https://img.shields.io/badge/version-2.7-blue)
![Platform](https://img.shields.io/badge/platform-Excel%20VBA-green)
![License](https://img.shields.io/badge/license-Proprietary-red)

---

## 🚀 Übersicht

Der Personalplaner ist eine vollständige Lösung zur Verwaltung von Mitarbeiterressourcen, Abwesenheiten und Auslastungsplanung mit einer intuitiven Custom Ribbon-Benutzeroberfläche.

### Hauptfeatures

- 📅 **Kalenderverwaltung** - Automatische Erstellung von Arbeitstageskalendern mit KWs, Feiertagen und Schulferien
- 👥 **Personalplanung** - Verwaltung von Mitarbeiterdaten und Abwesenheiten (Ferien, Krankheit, Militär, etc.)
- 📊 **Auslastungsberechnung** - Robuste UDFs für Berechnung verfügbarer Mitarbeiter und Auslastungsquoten
- 🎨 **Custom Ribbon UI** - Intuitive Bedienung über benutzerdefinierte Excel-Menüleiste
- 📑 **Wochenplan-Export** - Automatische Erstellung und Versand von KW-Plänen als PDF
- 🔍 **Filter & Projekte** - Filterbare Ansichten und Projektverwaltung
- 📈 **Dashboard** - Auswertungen und Visualisierungen

---

## 📋 Systemanforderungen

- Microsoft Excel 2010 oder neuer (empfohlen: Excel 2016+)
- Makros müssen aktiviert sein
- VBA7 oder kompatible Version

---

## 🎯 Schnellstart

1. Excel-Datei mit aktivierten Makros öffnen (.xlsm)
2. Custom Ribbon wird automatisch geladen
3. Navigation über Ribbon-Buttons:
   - **Heute** - Springt zum aktuellen Datum
   - **Übersicht** - Hauptansicht
   - **Auswertung** - Dashboard
   - **Filter** - Filterung aktivieren
   - **Projekt** - Projektverwaltung

---

## 📚 Dokumentation

Vollständige Informationen zu Features, Funktionen und technischen Details finden Sie in den **[Release Notes v2.7](RELEASE_NOTES_v2.7.md)**.

### Wichtige Abwesenheitscodes

| Code | Bedeutung |
|------|-----------|
| F | Ferien |
| Fx | Ferien nicht bewilligt |
| K | Krank |
| U | Unfall |
| WK | Militär |
| S | Schule |
| ÜK | Überbetrieblicher Kurs |
| T | Teilzeit |

---

## 🛠️ Technische Details

### Code-Struktur
- **15 VBA-Module** (*.bas, *.frm, *.doccls)
- **Custom Ribbon UI** mit IRibbonUI
- **ListObject-basierte** Datenverwaltung
- **Dictionary-optimierte** Lookup-Operationen

### Kernmodule
- `mKalender.bas` - Kalenderfunktionen
- `mBerechnung.bas` - Auslastungsberechnungen (UDFs)
- `mKWBlatt.bas` - Wochenplan-Export
- `CustomUI.bas` - Ribbon-Integration
- `UF_Filter.frm` / `UF_Projekte.frm` - UserForms

---

## 🔧 Wartung

### Zu pflegende Tabellen
- **Feiertage**: Name, Datum
- **Ferien**: Name, Start-Datum, End-Datum
- **Mitarbeiter**: Nummer, Name, Funktion, Team, Kontaktdaten

### Performance-Hinweise
- Berechnung ist auf "Manuell" gestellt (Performance-Optimierung)
- Manuelle Neuberechnung über Ribbon "Berechnen" oder `F9`
- Bei großen Datenmengen: Nicht benötigte Blätter ausblenden

---

## 📦 Release v2.7 (19.11.2025)

Vollständige Erstveröffentlichung mit komplettem Funktionsumfang.

**Was ist neu:**
- ✅ Kalenderverwaltung mit Arbeitstagen
- ✅ Robuste Auslastungsberechnungen
- ✅ Custom Ribbon UI
- ✅ Wochenplan-Export
- ✅ Filter & Projektverwaltung
- ✅ Dashboard mit Auswertungen
- ✅ Performance-Optimierungen

Siehe **[RELEASE_NOTES_v2.7.md](RELEASE_NOTES_v2.7.md)** für Details.

---

## 📄 Lizenz

Proprietär - Alle Rechte vorbehalten

---

## 🤝 Kontakt & Support

Bei Fragen zur Verwendung oder technischen Problemen wenden Sie sich bitte an den Systemadministrator.

---

**Entwickelt für effiziente Personalplanung und Ressourcenverwaltung**


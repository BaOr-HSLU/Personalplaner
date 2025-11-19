## Beschreibung

Beschreibe die Änderungen in diesem Pull Request.

## Art der Änderung

- [ ] 🐛 Bugfix (nicht-breaking change, der ein Problem behebt)
- [ ] ✨ Neues Feature (nicht-breaking change, der Funktionalität hinzufügt)
- [ ] 💥 Breaking Change (Fix oder Feature, das bestehende Funktionalität bricht)
- [ ] 📝 Dokumentation (nur Änderungen an Dokumentation)
- [ ] ♻️ Refactoring (Code-Änderung ohne Funktionsänderung)
- [ ] ⚡ Performance (Verbesserung der Performance)
- [ ] 🎨 Style (Formatierung, fehlende Semikolons, etc.)
- [ ] ✅ Tests (Hinzufügen oder Korrigieren von Tests)

## Verwandte Issues

Closes #(issue)
Fixes #(issue)
Relates to #(issue)

## Änderungen im Detail

### Neue Dateien
- `path/to/file.bas` - Beschreibung

### Geänderte Dateien
- `path/to/file.bas` - Beschreibung der Änderung

### Gelöschte Dateien
- `path/to/file.bas` - Grund für Löschung

## VBA-Module betroffen

- [ ] mKalender.bas
- [ ] mBerechnung.bas
- [ ] mAuslastung.bas
- [ ] mKWBlatt.bas
- [ ] mFilter.bas
- [ ] mFormatierung.bas
- [ ] mWertesammler.bas
- [ ] CustomUI.bas
- [ ] UserForms (UF_*)
- [ ] DieseArbeitsmappe.doccls
- [ ] Sonstiges: ___________

## Funktionsbereiche betroffen

- [ ] Kalenderverwaltung
- [ ] Personalplanung
- [ ] Auslastungsberechnung
- [ ] Ribbon UI
- [ ] Wochenplan-Export
- [ ] Filter
- [ ] Projektverwaltung
- [ ] Dashboard / Auswertungen
- [ ] Performance
- [ ] Dokumentation

## Screenshots (falls UI-Änderungen)

Falls UI-Änderungen vorgenommen wurden, bitte Screenshots hinzufügen:

**Vorher:**
<!-- Screenshot einfügen -->

**Nachher:**
<!-- Screenshot einfügen -->

## Test-Plan

Beschreibe, wie die Änderungen getestet wurden:

### Manuelle Tests durchgeführt

- [ ] Feature/Fix manuell getestet
- [ ] Edge Cases geprüft
- [ ] Regressionstest (bestehende Features funktionieren noch)
- [ ] Performance-Test (keine Verlangsamung)

### Spezifische Test-Schritte

1. Schritt 1
2. Schritt 2
3. Schritt 3

**Erwartetes Ergebnis:**
<!-- Beschreibung -->

**Tatsächliches Ergebnis:**
<!-- Beschreibung -->

## Checklist

### Code Quality

- [ ] Code folgt den Coding Standards (siehe CONTRIBUTING.md)
- [ ] `Option Explicit` in allen neuen Modulen
- [ ] Error Handling implementiert (`On Error GoTo`)
- [ ] Keine `Debug.Print` Statements im finalen Code
- [ ] Code-Kommentare hinzugefügt (@Description, @Param, @Return)
- [ ] Performance-Optimierungen berücksichtigt (ScreenUpdating, etc.)

### Dokumentation

- [ ] README.md aktualisiert (falls nötig)
- [ ] CHANGELOG.md aktualisiert
- [ ] Inline-Kommentare für komplexe Logik
- [ ] CONTRIBUTING.md gelesen und befolgt

### Testing

- [ ] Änderungen in Excel getestet
- [ ] Funktioniert in Excel 2016+
- [ ] Keine neuen VBA-Fehler eingeführt
- [ ] Bestehende Funktionalität nicht beeinträchtigt

### Git

- [ ] Branch ist aktuell mit `main`
- [ ] Commit-Messages folgen Conventional Commits
- [ ] Keine Merge-Konflikte
- [ ] Keine unnötigen Dateien committed

## Breaking Changes

Falls Breaking Changes vorhanden sind, beschreibe:

### Was bricht?

<!-- Beschreibung -->

### Migration Path

Wie können Nutzer ihre bestehenden Setups anpassen?

1. Schritt 1
2. Schritt 2

## Performance-Auswirkungen

- [ ] Keine Performance-Auswirkungen
- [ ] Performance-Verbesserung
- [ ] Potenzielle Performance-Verschlechterung (beschreiben)

**Details:**
<!-- Falls Performance-Änderungen, bitte beschreiben -->

## Abhängigkeiten

Neue Abhängigkeiten oder geänderte Systemanforderungen?

- [ ] Keine neuen Abhängigkeiten
- [ ] Neue VBA-Referenzen erforderlich: ___________
- [ ] Höhere Excel-Version erforderlich: ___________

## Zusätzliche Notizen

Weitere Informationen für Reviewer:

<!-- Zusätzlicher Kontext, Designentscheidungen, offene Fragen, etc. -->

## Reviewer-Hinweise

Worauf sollten Reviewer besonders achten?

- [ ] Logik in Funktion X
- [ ] Performance bei großen Datenmengen
- [ ] UI/UX-Änderungen
- [ ] Sonstiges: ___________

---

**Bereit für Review** ✅

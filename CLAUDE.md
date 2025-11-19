# Arbeiten mit Claude im Personalplaner-Projekt

Diese Anleitung beschreibt Best Practices für die Zusammenarbeit mit Claude (AI) bei der Entwicklung des Personalplaners.

## 📋 Inhaltsverzeichnis

- [Issue-Tracking](#issue-tracking)
- [Template-Verwendung](#template-verwendung)
- [Workflow mit Claude](#workflow-mit-claude)
- [Claude-spezifische Konventionen](#claude-spezifische-konventionen)
- [Kommunikations-Best-Practices](#kommunikations-best-practices)
- [Beispiel-Workflows](#beispiel-workflows)

---

## Issue-Tracking

### ✅ Pflicht: Issues verwenden

**WICHTIG:** Alle Arbeiten am Projekt **MÜSSEN** über GitHub Issues getrackt werden.

#### Warum?
- 📊 Nachvollziehbarkeit aller Änderungen
- 🔍 Suchbare Historie
- 👥 Team-Transparenz
- 🔗 Verknüpfung von Commits/PRs mit Issues
- 📈 Projekt-Fortschritt tracking

### Issue erstellen (PFLICHT vor jeder Arbeit)

#### 1. Via GitHub Web UI

```
https://github.com/BaOr-HSLU/Personalplaner/issues/new/choose
```

**Wähle das passende Template:**
- 🐛 **Bug Report** - Für Fehler und Bugs
- ✨ **Feature Request** - Für neue Features
- ❓ **Question** - Für Fragen

#### 2. Via gh CLI (falls verfügbar)

```bash
# Bug Report
gh issue create --template bug_report.md --title "[BUG] Beschreibung"

# Feature Request
gh issue create --template feature_request.md --title "[FEATURE] Beschreibung"

# Question
gh issue create --template question.md --title "[FRAGE] Beschreibung"
```

#### 3. Claude bitten, Issue zu erstellen

```
Claude, bitte erstelle ein Issue für:
- Bug: Fehler bei Datumsberechnung in mKalender
- Feature: CSV-Import für Feiertage
- Frage: Wie funktioniert die Ribbon-Aktualisierung?
```

**Claude wird dann:**
1. Template ausfüllen
2. Issue-Nummer bereitstellen
3. Diese in Commits referenzieren

---

## Template-Verwendung

### GitHub Issue Templates (PFLICHT)

Alle Issues **MÜSSEN** eines der Templates verwenden:

#### Bug Report Template
```markdown
## Beschreibung des Problems
[Klare Beschreibung]

## Schritte zur Reproduktion
1. Gehe zu '...'
2. Klicke auf '...'
3. Fehler tritt auf

## Erwartetes Verhalten
[Was sollte passieren]

## Tatsächliches Verhalten
[Was passiert stattdessen]

## Umgebung
- Excel-Version: [z.B. Excel 2016]
- Betriebssystem: [z.B. Windows 10]
- Personalplaner-Version: [z.B. v2.7.0]

## Betroffenes Modul
- [ ] Kalenderverwaltung
- [ ] Personalplanung
- [x] Auslastungsberechnung
```

#### Feature Request Template
```markdown
## Feature-Beschreibung
[Klare Beschreibung des gewünschten Features]

## Problem / Motivation
[Welches Problem würde dieses Feature lösen?]

## Vorgeschlagene Lösung
[Wie sollte das Feature funktionieren?]

## Betroffene Bereiche
- [x] Kalenderverwaltung
- [ ] Ribbon UI
```

#### Question Template
```markdown
## Deine Frage
[Stelle deine Frage klar und präzise]

## Kontext
[Beschreibe den Kontext]

## Was hast du bereits versucht?
- [x] README.md gelesen
- [x] Code-Kommentare angeschaut
```

### Pull Request Template (PFLICHT)

Jeder PR **MUSS** das Template verwenden:

```markdown
## Beschreibung
[Beschreibe die Änderungen]

## Art der Änderung
- [ ] 🐛 Bugfix
- [x] ✨ Neues Feature
- [ ] 💥 Breaking Change

## Verwandte Issues
Closes #42
Fixes #38

## Checklist
- [x] Code folgt Coding Standards
- [x] CHANGELOG.md aktualisiert
- [x] Tests durchgeführt
```

---

## Workflow mit Claude

### Standard-Workflow

#### 1. Issue erstellen (vor jeder Arbeit)

**Du sagst zu Claude:**
```
Erstelle ein Issue für: [Beschreibung]
```

**Claude erstellt:**
- Issue mit korrektem Template
- Issue-Nummer (z.B. #42)
- Labels (bug, enhancement, question)

#### 2. Branch erstellen

**Du sagst zu Claude:**
```
Erstelle einen Feature-Branch für Issue #42
```

**Claude führt aus:**
```bash
git checkout main
git pull origin main
git checkout -b feature/issue-42-beschreibung
```

#### 3. Entwicklung

**Du sagst zu Claude:**
```
Implementiere die Lösung für Issue #42
```

**Claude:**
- Implementiert Code
- Erstellt Tests
- Dokumentiert Änderungen
- Aktualisiert CHANGELOG.md

#### 4. Commit mit Issue-Referenz (PFLICHT)

**Claude committet automatisch mit Issue-Referenz:**
```bash
git commit -m "feat(kalender): Füge CSV-Import hinzu

- CSV-Parser implementiert
- Validierung für Datumsformate
- Error-Handling hinzugefügt

Relates to #42"
```

**Wichtige Keywords für GitHub:**
- `Closes #42` - Schließt Issue beim Merge
- `Fixes #42` - Behebt Issue beim Merge
- `Resolves #42` - Löst Issue beim Merge
- `Relates to #42` - Verknüpft mit Issue (schließt nicht)

#### 5. Pull Request erstellen

**Du sagst zu Claude:**
```
Erstelle einen PR für diesen Branch
```

**Claude:**
- Füllt PR-Template aus
- Referenziert Issues
- Checklist abgehakt
- Test-Plan beschrieben

#### 6. Nach Merge: Issue schließen

GitHub schließt Issues automatisch wenn im PR steht:
```
Closes #42
```

---

## Claude-spezifische Konventionen

### Issue-Erstellung durch Claude

**Best Practice:**

```
Claude, erstelle ein Issue:

Typ: Bug / Feature / Question
Titel: [Kurze Beschreibung]
Beschreibung: [Details]
Betroffene Module: [Liste]
```

**Claude antwortet:**
```
✅ Issue erstellt: #42
📝 Titel: [BUG] Datumsberechnung in mKalender
🔗 https://github.com/BaOr-HSLU/Personalplaner/issues/42
```

### Commits referenzieren Issues

**Claude verwendet automatisch:**

```bash
# Feature
git commit -m "feat(scope): Beschreibung

Details...

Relates to #42"

# Bugfix (schließt Issue)
git commit -m "fix(scope): Beschreibung

Details...

Fixes #42"

# Dokumentation
git commit -m "docs: Update README

Relates to #42"
```

### CHANGELOG aktualisieren

**Claude fügt automatisch zu CHANGELOG.md hinzu:**

```markdown
## [Unreleased]

### Hinzugefügt
- CSV-Import für Feiertage (#42)

### Behoben
- Datumsberechnung bei Schaltjahren (#38)
```

---

## Kommunikations-Best-Practices

### ✅ Gute Anfragen an Claude

```
✅ "Erstelle ein Feature-Request-Issue für CSV-Import bei Feiertagen"
✅ "Implementiere Lösung für Issue #42 gemäß CONTRIBUTING.md"
✅ "Erstelle PR für Branch feature/csv-import mit allen Checklists"
✅ "Aktualisiere CHANGELOG.md für Version 2.8.0"
✅ "Fixe Bug #38 und referenziere das Issue im Commit"
```

### ❌ Schlechte Anfragen

```
❌ "Mach irgendwas mit dem Kalender" (zu vage)
❌ "Füge Feature hinzu" (kein Issue-Tracking)
❌ "Committe das" (keine Issue-Referenz)
❌ "Fix den Bug" (welches Issue?)
```

### Issue-First Approach (PFLICHT)

**IMMER in dieser Reihenfolge:**

1. ✅ **Issue erstellen** (auch für kleine Änderungen!)
2. ✅ **Branch erstellen** (feature/issue-XX)
3. ✅ **Implementieren**
4. ✅ **Commit mit Issue-Referenz**
5. ✅ **PR mit Issue-Verknüpfung**
6. ✅ **Merge → Issue wird automatisch geschlossen**

**NIEMALS:**
❌ Direkt committen ohne Issue
❌ PR ohne Issue-Referenz
❌ Commits ohne Issue-Nummer

---

## Beispiel-Workflows

### Workflow 1: Neues Feature

```
👤 Du:
"Claude, ich möchte einen CSV-Import für Feiertage hinzufügen"

🤖 Claude:
"Ich erstelle zuerst ein Feature-Request-Issue..."
✅ Issue #45 erstellt
"Soll ich mit der Implementierung beginnen?"

👤 Du:
"Ja, bitte implementiere das Feature"

🤖 Claude:
1. Erstellt Branch: feature/issue-45-csv-import
2. Implementiert Code in mKalender.bas
3. Aktualisiert CHANGELOG.md
4. Committet mit: "feat(kalender): CSV-Import für Feiertage

   - CSV-Parser implementiert
   - Validierung hinzugefügt
   - Tests durchgeführt

   Relates to #45"
5. Pushed Branch
6. Erstellt PR mit Template
7. Verknüpft PR mit Issue #45

"✅ PR erstellt, bereit für Review"
```

### Workflow 2: Bugfix

```
👤 Du:
"Claude, die Datumsberechnung in mKalender.bas funktioniert nicht bei Schaltjahren"

🤖 Claude:
"Ich erstelle ein Bug-Report-Issue..."

## Issue #46: [BUG] Datumsberechnung bei Schaltjahren

**Beschreibung:** Schaltjahre werden nicht korrekt behandelt
**Betroffenes Modul:** mKalender.bas
**Funktion:** FindeDatumsspalte()

✅ Issue erstellt
"Soll ich den Bug fixen?"

👤 Du:
"Ja, fixe den Bug"

🤖 Claude:
1. Branch: bugfix/issue-46-schaltjahr
2. Analysiert Code
3. Implementiert Fix
4. Testet Edge Cases
5. Committet: "fix(kalender): Schaltjahr-Behandlung korrigiert

   - Funktion FindeDatumsspalte() berücksichtigt jetzt Schaltjahre
   - Test mit 29.02.2024 hinzugefügt

   Fixes #46"
6. Erstellt PR
7. PR merged → Issue #46 automatisch geschlossen
```

### Workflow 3: Dokumentation

```
👤 Du:
"Claude, erstelle eine FAQ-Sektion in der README"

🤖 Claude:
"Soll ich dafür ein Issue erstellen?"

👤 Du:
"Ja"

🤖 Claude:
✅ Issue #47: [DOCS] FAQ-Sektion in README

1. Branch: docs/issue-47-faq
2. Fügt FAQ zu README.md hinzu
3. Committet: "docs: FAQ-Sektion hinzugefügt

   - Häufige Fragen zu Installation
   - Troubleshooting-Tipps
   - Links zu detaillierter Doku

   Closes #47"
4. PR erstellt und gemerged
```

### Workflow 4: Multiple Issues

```
👤 Du:
"Claude, ich brauche:
1. CSV-Import für Feiertage
2. PDF-Export verbessern
3. Ribbon-Button hinzufügen"

🤖 Claude:
"Ich erstelle für jede Aufgabe ein separates Issue..."

✅ Issue #48: [FEATURE] CSV-Import für Feiertage
✅ Issue #49: [ENHANCEMENT] PDF-Export verbessern
✅ Issue #50: [FEATURE] Ribbon-Button für Schnellexport

"In welcher Reihenfolge soll ich diese implementieren?"

👤 Du:
"Beginne mit #48"

🤖 Claude:
[Arbeitet an Issue #48, dann #49, dann #50]
[Jeweils eigener Branch, eigener PR, eigene Issue-Referenz]
```

---

## Tracking & Reporting

### Status-Abfrage

```
👤 Du:
"Claude, zeige offene Issues"

🤖 Claude:
📊 Offene Issues:
- #48: [FEATURE] CSV-Import (in progress, branch: feature/issue-48)
- #49: [ENHANCEMENT] PDF-Export (pending)
- #50: [FEATURE] Ribbon-Button (pending)

✅ Geschlossene Issues:
- #45: CSV-Import für Feiertage (merged, v2.8.0)
- #46: Schaltjahr-Bug (merged, v2.7.1)
```

### Changelog generieren

```
👤 Du:
"Claude, generiere Changelog für v2.8.0"

🤖 Claude:
```markdown
## [2.8.0] - 2025-11-20

### Hinzugefügt
- CSV-Import für Feiertage (#45)
- Ribbon-Button für Schnellexport (#50)

### Geändert
- PDF-Export Performance verbessert (#49)

### Behoben
- Schaltjahr-Behandlung in Datumsberechnung (#46)
```

---

## Checklists für Claude

### ✅ Vor jeder Arbeit

- [ ] Issue existiert (oder erstellen)
- [ ] Issue-Nummer bekannt
- [ ] Template verwendet
- [ ] Labels gesetzt

### ✅ Bei Implementierung

- [ ] Branch von main erstellt
- [ ] Code nach Standards (CONTRIBUTING.md)
- [ ] CHANGELOG.md aktualisiert
- [ ] Tests durchgeführt

### ✅ Bei Commit

- [ ] Conventional Commit Format
- [ ] Issue-Referenz im Commit-Body
- [ ] Beschreibende Commit-Message
- [ ] Korrekte Keywords (Closes/Fixes/Relates)

### ✅ Bei Pull Request

- [ ] PR-Template vollständig ausgefüllt
- [ ] Issue verknüpft (Closes #XX)
- [ ] Checklist abgehakt
- [ ] Test-Plan beschrieben
- [ ] CHANGELOG.md aktualisiert

---

## Tools & Commands

### GitHub CLI Integration (falls verfügbar)

```bash
# Issue listen
gh issue list

# Issue erstellen
gh issue create --template bug_report.md

# Issue Details
gh issue view 42

# PR erstellen
gh pr create --fill

# PR Status
gh pr status
```

### Git Aliases (empfohlen)

```bash
# In ~/.gitconfig oder .git/config

[alias]
    # Issue-bezogene Commits
    ci = "!f() { git commit -m \"$1\n\nRelates to #$2\"; }; f"
    fix = "!f() { git commit -m \"fix: $1\n\nFixes #$2\"; }; f"
    feat = "!f() { git commit -m \"feat: $1\n\nRelates to #$2\"; }; f"

    # Branch für Issue
    issue-branch = "!f() { git checkout -b feature/issue-$1-${2}; }; f"
```

**Verwendung:**
```bash
git issue-branch 42 csv-import  # erstellt feature/issue-42-csv-import
git feat "CSV Import hinzugefügt" 42  # committet mit Issue-Referenz
```

---

## Fehler vermeiden

### ❌ NICHT tun

```bash
# Ohne Issue arbeiten
git commit -m "fix stuff"  # ❌ Keine Issue-Referenz

# Direkt auf main pushen
git push origin main  # ❌ Immer über PR!

# PR ohne Template
[Leere PR-Beschreibung]  # ❌ Template verwenden!

# Issue nicht verlinken
git commit -m "feat: neues Feature"  # ❌ Wo ist das Issue?
```

### ✅ IMMER tun

```bash
# Mit Issue-Referenz
git commit -m "feat(scope): Beschreibung

Details...

Relates to #42"  # ✅

# Über PR
feature-branch → PR → main  # ✅

# Template verwenden
[PR-Template vollständig ausgefüllt]  # ✅

# Issue verlinken
Closes #42 im PR  # ✅
```

---

## Zusammenfassung

### Goldene Regeln

1. **📝 Jede Arbeit = Ein Issue**
2. **📋 Immer Templates verwenden**
3. **🔗 Commits referenzieren Issues**
4. **✅ PRs schließen Issues automatisch**
5. **📊 CHANGELOG aktualisieren**
6. **🔄 Issue-First Workflow**

### Kommunikation mit Claude

```
Format: "Claude, [Aktion] für Issue #XX"

Beispiele:
✅ "Claude, erstelle Issue für CSV-Import"
✅ "Claude, implementiere Lösung für Issue #42"
✅ "Claude, erstelle PR für Issue #42"
✅ "Claude, aktualisiere CHANGELOG für v2.8.0"
```

---

## Weitere Ressourcen

- [CONTRIBUTING.md](CONTRIBUTING.md) - Contribution Guidelines
- [DEVELOPMENT.md](DEVELOPMENT.md) - Developer Documentation
- [CHANGELOG.md](CHANGELOG.md) - Version History
- [GitHub Issues](https://github.com/BaOr-HSLU/Personalplaner/issues)

---

**Mit diesem Workflow bleibt alles nachvollziehbar und sauber dokumentiert!** 🚀

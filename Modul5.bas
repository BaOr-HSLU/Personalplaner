Attribute VB_Name = "Modul5"
Option Explicit

'==============================================================================
' Prozedur: WR_Erinnerung_Versenden
' Zweck: Sendet EINE E-Mail an alle Mitarbeiter, die noch keinen
'        Wochenrapport abgegeben haben
' Voraussetzung: E-Mail-Adresse muss in der 3. Zeile der Namenszelle stehen
'                (getrennt durch vbNewLine)
' Autor: [Name]
' Datum: [Datum]
' Version: 1.0
'==============================================================================
Public Sub WR_Anfordern()
    '==========================================================================
    ' Deklaration und Initialisierung der Variablen
    '==========================================================================
    
    ' Arbeitsblätter und Outlook-Objekte
    Dim wsWochenplan As Worksheet
    Dim outlookApp As Object
    Dim mailItem As Object
    
    ' Listen und Bereiche
    Dim lo As ListObject
    Dim rng As Range
    Dim cell As Range
    
    ' Mitarbeiter-Daten
    Dim MAB As Scripting.Dictionary
    Dim lkey As Variant
    Dim MABRow As Long
    Dim MABName As String
    Dim MABNameFull As String
    Dim MABEmail As String
    Dim nameArray() As String
    
    ' E-Mail-Listen und Inhalt
    Dim emailList As String
    Dim emailBetreff As String
    Dim emailText As String
    Dim KW As String
    
    ' Zähler und Status
    Dim countEmails As Long
    Dim countErrors As Long
    Dim errorList As String
    
    ' Originaleinstellungen speichern
    Dim originalScreenUpdating As Boolean
    originalScreenUpdating = Application.ScreenUpdating
    
    '==========================================================================
    ' Fehlerbehandlung und Performance-Optimierung
    '==========================================================================
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    '==========================================================================
    ' Arbeitsblatt und Daten initialisieren
    '==========================================================================
    Set wsWochenplan = ActiveSheet
    
    ' Kalenderwoche für E-Mail-Betreff auslesen
    KW = wsWochenplan.Range("A3").Value
    
    '==========================================================================
    ' Eindeutige Mitarbeiter aus Wochenplan sammeln (Spalte A)
    '==========================================================================
    Set lo = wsWochenplan.Range("E7").ListObject
    Set rng = Intersect(lo.DataBodyRange, wsWochenplan.Range("A:A"))
    
    ' Sammle alle eindeutigen Mitarbeiter
    Set MAB = SammleEindeutigeWerteSchnellRng(rng, includeHidden:=False)
    
    '==========================================================================
    ' Zähler initialisieren
    '==========================================================================
    emailList = ""
    countEmails = 0
    countErrors = 0
    errorList = ""
    
    '==========================================================================
    ' Alle E-Mail-Adressen der Mitarbeiter sammeln
    '==========================================================================
    For Each lkey In MAB.Keys
        On Error Resume Next
        
        ' Mitarbeiterzeile finden
        Set cell = rng.Find(lkey)
        If Not cell Is Nothing Then
            MABRow = cell.row
            
            ' Prüfen ob Mitarbeiter NICHT ausgelassen werden soll (Spalte K)
            ' Nur wenn Spalte K = FALSE, dann E-Mail hinzufügen
            If Not wsWochenplan.Cells(MABRow, 11).Value Then
                
                '--------------------------------------------------------------
                ' Mitarbeiterdaten extrahieren
                '--------------------------------------------------------------
                ' Vollständiger Zelleninhalt aus Spalte B (Name mit möglichen Zeilen)
                MABNameFull = wsWochenplan.Cells(MABRow, 2).Value
                
                ' Nach Zeilenumbrüchen aufteilen
                nameArray = Split(MABNameFull, vbNewLine)
                
                ' Erste Zeile = Name für Fehlerprotokoll
                MABName = nameArray(0)
                
                ' Dritte Zeile = E-Mail-Adresse (Index 2, da bei 0 beginnend)
                If UBound(nameArray) >= 2 Then
                    MABEmail = Trim(nameArray(2))
                    
                    ' Validierung: Prüfen ob E-Mail-Adresse vorhanden und gültig
                    If MABEmail <> "" And InStr(MABEmail, "@") > 0 Then
                        
                        '----------------------------------------------
                        ' E-Mail-Adresse zur Liste hinzufügen
                        '----------------------------------------------
                        If emailList <> "" Then
                            emailList = emailList & "; "
                        End If
                        emailList = emailList & MABEmail
                        countEmails = countEmails + 1
                        
                    Else
                        ' E-Mail-Adresse fehlt oder ist ungültig
                        countErrors = countErrors + 1
                        errorList = errorList & "- " & MABName & ": Keine gültige E-Mail-Adresse" & vbNewLine
                    End If
                Else
                    ' Weniger als 3 Zeilen in der Namenszelle
                    countErrors = countErrors + 1
                    errorList = errorList & "- " & MABName & ": E-Mail-Adresse fehlt (weniger als 3 Zeilen)" & vbNewLine
                End If
            End If
        End If
        
        On Error GoTo ErrorHandler
    Next lkey
    
    '==========================================================================
    ' Prüfen ob E-Mail-Adressen gefunden wurden
    '==========================================================================
    If emailList = "" Then
        MsgBox "Keine gültigen E-Mail-Adressen gefunden." & vbNewLine & vbNewLine & _
               "Bitte prüfen Sie, ob die E-Mail-Adressen in der 3. Zeile der Namenszellen stehen.", _
               vbExclamation, "Keine Empfänger"
        GoTo CleanupAndExit
    End If
    
    '==========================================================================
    ' Outlook-Anwendung initialisieren
    '==========================================================================
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    
    ' Falls Outlook nicht läuft, neu starten
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler
    
    ' Prüfen ob Outlook erfolgreich initialisiert wurde
    If outlookApp Is Nothing Then
        MsgBox "Outlook konnte nicht gestartet werden." & vbNewLine & _
               "Bitte stellen Sie sicher, dass Microsoft Outlook installiert ist.", _
               vbCritical, "Outlook-Fehler"
        GoTo CleanupAndExit
    End If
    
    '==========================================================================
    ' E-Mail-Inhalt definieren
    '==========================================================================
    emailBetreff = "Erinnerung: Wochenrapport " & KW & " abgeben"
    emailText = "Hallo zusammen," & vbNewLine & vbNewLine & _
                "bitte gebt noch euren Wochenrapport ab." & vbNewLine & vbNewLine & _
                "Vielen Dank!" & vbNewLine & _
                "Mit freundlichen Grüssen"
    
    '==========================================================================
    ' Eine E-Mail an alle Mitarbeiter erstellen
    '==========================================================================
    Set mailItem = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    With mailItem
        ' Alle E-Mail-Adressen im TO-Feld
        .To = emailList
        
        ' Optional: Blind Copy (BCC) verwenden für Datenschutz
        ' .BCC = emailList
        ' .To = "" ' TO-Feld leer lassen bei BCC
        
        .Subject = emailBetreff
        .Body = emailText
        
        ' Optional: Wichtigkeit setzen
        .Importance = 2 ' 2 = olImportanceHigh
        
        ' E-Mail anzeigen (zur Kontrolle) oder direkt senden
        ' Zum Testen: .Display verwenden
        ' Für automatischen Versand: .Send verwenden
        
        .Display ' Zum Testen - E-Mail wird angezeigt
        ' .Send   ' Für automatischen Versand - auskommentieren
    End With
    
    '==========================================================================
    ' Aufräumen und Abschluss
    '==========================================================================
CleanupAndExit:
    
    ' Originaleinstellungen wiederherstellen
    Application.ScreenUpdating = originalScreenUpdating
    
    ' Erfolgsmeldung zusammenstellen
    If countEmails > 0 Then
        Dim successMessage As String
        successMessage = "E-Mail an " & countEmails & " Empfänger wurde erstellt."
        
        If countErrors > 0 Then
            successMessage = successMessage & vbNewLine & vbNewLine & _
                            countErrors & " Fehler aufgetreten:" & vbNewLine & vbNewLine & _
                            errorList
            MsgBox successMessage, vbExclamation, "E-Mail erstellt mit Fehlern"
        Else
            MsgBox successMessage, vbInformation, "E-Mail erfolgreich erstellt"
        End If
    End If
    
    ' Objekte freigeben
    Set mailItem = Nothing
    Set outlookApp = Nothing
    Set MAB = Nothing
    Set cell = Nothing
    Set rng = Nothing
    Set lo = Nothing
    Set wsWochenplan = Nothing
    
    Exit Sub
    
    '==========================================================================
    ' Fehlerbehandlung
    '==========================================================================
ErrorHandler:
    Application.ScreenUpdating = originalScreenUpdating
    
    MsgBox "Fehler " & Err.Number & " ist aufgetreten:" & vbNewLine & vbNewLine & _
           Err.Description & vbNewLine & vbNewLine & _
           "Quelle: " & Err.Source, _
           vbCritical, "Fehler beim E-Mail-Versand"
    
    ' Fehler protokollieren
    Debug.Print "Fehler in WR_Erinnerung_Versenden: " & Err.Number & " - " & Err.Description
    
    ' Objekte freigeben
    Set mailItem = Nothing
    Set outlookApp = Nothing
    
End Sub


'==============================================================================
' Prozedur: WR_Erstellen
' Zweck: Erstellt automatisch Wochenrapporte für alle Mitarbeiter basierend
'        auf den Daten aus dem aktiven Wochenplan
' Autor: [Name]
' Datum: [Datum]
' Version: 2.0 - Optimiert
'==============================================================================
Public Sub WR_Erstellen()
    '==========================================================================
    ' Deklaration und Initialisierung der Variablen
    '==========================================================================
    
    ' Arbeitsblätter
    Dim wsTemplate As Worksheet
    Dim wsWochenplan As Worksheet
    Dim wsMAB As Worksheet
    Dim newWB As Workbook
    
    ' Wochenplan-Daten
    Dim KW As String
    Dim startdate As Date
    Dim enddate As Date
    
    ' Dictionaries für Projekte und Mitarbeiter
    Dim PROJEKTE As Scripting.Dictionary
    Dim MAB As Scripting.Dictionary
    
    ' Hilfsvariablen
    Dim lo As ListObject
    Dim rng As Range
    Dim cell As Range
    Dim lkey As Variant
    Dim MABRow As Long
    Dim MABName As String
    Dim MABProjrow As Long
    Dim cntRapporte As Long
    Dim lpath As String
    Dim row As Long
    Dim Kom As String
    Dim Bem As String
    Dim larr() As String
    Dim carr() As String
    Dim strComment As String
    Dim strProj As String
    Dim i As Long
    Dim MABProj As Range
    
    ' Originaleinstellungen speichern für spätere Wiederherstellung
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEnableEvents As Boolean
    Dim originalDisplayAlerts As Boolean
    
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalEnableEvents = Application.EnableEvents
    originalDisplayAlerts = Application.DisplayAlerts
    
    '==========================================================================
    ' Fehlerbehandlung und Performance-Optimierung aktivieren
    '==========================================================================
    On Error GoTo ErrorHandler
    
    ' Performance-Optimierung: Display-Updates und Events deaktivieren
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    '==========================================================================
    ' Arbeitsblätter zuweisen
    '==========================================================================
    Set wsTemplate = shWRTemplate
    Set wsWochenplan = ActiveSheet
    wsTemplate.Visible = xlSheetVisible
    
    '==========================================================================
    ' Wochenplan-Daten auslesen
    '==========================================================================
    KW = wsWochenplan.Range("A3").Value
    startdate = wsWochenplan.Range("E4").Value
    enddate = wsWochenplan.Range("F4").Value
    
    '==========================================================================
    ' Schritt 1: Eindeutige Projekte aus Wochenplan sammeln (Spalten E:I)
    '==========================================================================
    Set lo = wsWochenplan.Range("E7").ListObject
    Set rng = Intersect(lo.DataBodyRange, wsWochenplan.Range("E:I"))
    
    ' Sammle alle eindeutigen Projektwerte
    Set PROJEKTE = SammleEindeutigeWerteSchnellRng(rng, includeHidden:=False, OnlyFirstLine:=True)
    
    '==========================================================================
    ' Schritt 2: Projekte filtern - Bekannte Projekte aus Liste entfernen
    ' Diese müssen nicht manuell erfasst werden
    '==========================================================================
    Dim lrng As Range
    Set lrng = Tabelle5.ListObjects("Tabelle6").DataBodyRange.Columns(2)
    
    For Each cell In lrng.Cells
        Dim arr As Variant
        arr = Split(cell.Value, Chr(10))
        
        If UBound(arr) > 0 Then
            ' Array hat mehr als 1 Element = mehrere Zeilen
            lkey = arr(0)
        Else
            ' Array hat nur 1 Element = keine Zeilenumbrüche
            lkey = cell.Value2
        End If
        
        If PROJEKTE.Exists(lkey) Then
            PROJEKTE.Remove lkey
        End If
    Next cell

    
    '==========================================================================
    ' Schritt 3: Neue Projekte verarbeiten - Kommissionsnummer und Bemerkung
    ' erfassen oder aus bestehenden Daten laden
    '==========================================================================
    For Each lkey In PROJEKTE.Keys
        ' Prüfen ob Projekt bereits in der Projektliste vorhanden ist
        If wsProjekte.UsedRange.Resize(, 1).Find(lkey) Is Nothing Then
            '--------------------------------------------------------------
            ' Neues Projekt: Benutzereingaben erfragen
            '--------------------------------------------------------------
            row = wsProjekte.UsedRange.Rows.Count + 1
            
InputsPrompt:
            Kom = Application.InputBox("Kommissionsnummer für " & lkey, "Kommissionsnummer", Type:=2)
            
            ' Prüfen ob Benutzer abgebrochen hat
            If Kom = "False" Then GoTo CleanupAndExit
            
            Bem = Application.InputBox("Bemerkung für " & lkey, "Bemerkung", Type:=2)
            
            ' Prüfen ob Benutzer abgebrochen hat
            If Bem = "False" Then GoTo CleanupAndExit
            
            ' Bestätigung einholen
            If MsgBox("Soll das Projekt gespeichert werden?" & vbNewLine & vbNewLine & _
                      "Projekt: " & lkey & vbNewLine & _
                      "Kommission: " & Kom & vbNewLine & _
                      "Bemerkung: " & Bem, _
                      vbYesNo + vbQuestion, "Projekt speichern?") = vbYes Then
                
                ' Projekt in Projektliste speichern
                wsProjekte.Cells(row, 1).Value = lkey
                wsProjekte.Cells(row, 2).Value = Kom
                wsProjekte.Cells(row, 3).Value = Bem
            End If
            
            ' Projekt-Daten im Dictionary speichern
            PROJEKTE(lkey) = Kom & ";" & Bem
            
        Else
            '--------------------------------------------------------------
            ' Bestehendes Projekt: Daten aus Projektliste laden
            '--------------------------------------------------------------
            row = wsProjekte.UsedRange.Resize(, 1).Find(lkey).row
            Kom = wsProjekte.Cells(row, 2).Value
            Bem = wsProjekte.Cells(row, 3).Value
            
            ' Benutzer fragen ob bestehende Daten verwendet werden sollen
            If MsgBox("Sollen die Projektdaten geladen werden?" & vbNewLine & vbNewLine & _
                      "Projekt: " & lkey & vbNewLine & _
                      "Kommission: " & Kom & vbNewLine & _
                      "Bemerkung: " & Bem, _
                      vbYesNo + vbQuestion, "Daten laden") = vbNo Then
                ' Wenn nein, Eingabe erneut erfragen
                GoTo InputsPrompt
            End If
            
            ' Projekt-Daten im Dictionary speichern
            PROJEKTE(lkey) = Kom & ";" & Bem
        End If
    Next lkey
    
    '==========================================================================
    ' Schritt 4: Eindeutige Mitarbeiter aus Wochenplan sammeln (Spalte A)
    '==========================================================================
    Set rng = Intersect(lo.DataBodyRange, wsWochenplan.Range("A:A"))
    Set MAB = SammleEindeutigeWerteSchnellRng(rng, includeHidden:=False)
    
    '==========================================================================
    ' Schritt 5: Neue Arbeitsmappe für Wochenrapporte erstellen
    '==========================================================================
    lpath = ActiveWorkbook.Path
    Set newWB = Workbooks.Add
    
    ' Arbeitsmappe speichern
    newWB.SaveAs lpath & "\Wochenrapporte_" & KW & ".xlsm", xlOpenXMLWorkbookMacroEnabled
    
    ' Performance-Einstellungen auch für neue Arbeitsmappe setzen
    With newWB.Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    ' Zähler für erstellte Rapporte initialisieren
    cntRapporte = 0
    
    '==========================================================================
    ' Schritt 6: Für jeden Mitarbeiter einen Wochenrapport erstellen
    '==========================================================================
    For Each lkey In MAB.Keys
        On Error Resume Next ' Fehler temporär ignorieren für MAB-Übersprung
        
        ' Mitarbeiterzeile finden
        Set cell = rng.Find(lkey)
        If cell Is Nothing Then GoTo SkipMAB
        
        MABRow = cell.row
        
        ' Prüfen ob Mitarbeiter nicht ausgelassen werden soll (Spalte K)
        If Not wsWochenplan.Cells(MABRow, 11).Value Then
            
            '--------------------------------------------------------------
            ' Mitarbeitername extrahieren (erste Zeile bei mehrzeiligen Namen)
            '--------------------------------------------------------------
            MABName = Split(wsWochenplan.Cells(MABRow, 2).Value, vbLf)(0)
            
            '--------------------------------------------------------------
            ' Template kopieren und als neues Blatt einfügen
            '--------------------------------------------------------------
            wsTemplate.Copy After:=newWB.Sheets(newWB.Sheets.Count)
            Set wsMAB = newWB.Sheets(newWB.Sheets.Count)
            cntRapporte = cntRapporte + 1
            
            ' Arbeitsblatt benennen
            wsMAB.Name = MABName
            
            '--------------------------------------------------------------
            ' Kopfzeilen des Wochenrapports ausfüllen
            '--------------------------------------------------------------
            wsMAB.Range("A2").Value = "Wochenrapport von: " & MABName
            wsMAB.Range("E2").Value = "Datum von: " & Format(startdate, "DD.MM.YYYY")
            wsMAB.Range("J2").Value = "bis: " & Format(enddate, "DD.MM.YYYY")
            wsMAB.Range("N2").Value = "Kalenderwoche: " & Right(KW, 2)
            
CheckPointPROJEKTE:
            '--------------------------------------------------------------
            ' Projekte und Stunden für jeden Wochentag eintragen
            '--------------------------------------------------------------
            i = 0 ' Tag-Zähler
            Set MABProj = wsMAB.UsedRange.Resize(, 1).offset(0, 13) ' Spalte N (Projekt-Spalte)
            
            ' Durch alle Wochentage iterieren (Spalten E:I = Mo-Fr)
            For Each cell In wsWochenplan.Range(wsWochenplan.Cells(MABRow, 5), _
                                                 wsWochenplan.Cells(MABRow, 9)).Cells
                i = i + 1 ' Tag-Index (1=Montag, 5=Freitag)
                
                ' Spezielle Abwesenheiten behandeln
                Select Case cell.Value
                    Case "Krank"
                        MABProjrow = 29
                        wsMAB.Cells(MABProjrow, i + 2).Value = 8
                        
                    Case "Unfall"
                        MABProjrow = 28
                        wsMAB.Cells(MABProjrow, i + 2).Value = 8
                        
                    Case "Militär"
                        MABProjrow = 27
                        wsMAB.Cells(MABProjrow, i + 2).Value = 8
                        
                    Case "Ferien"
                        MABProjrow = 26
                        wsMAB.Cells(MABProjrow, i + 2).Value = 8
                        
                    Case "Schule", "Überbetr.Kurs"
                        ' Diese werden im Rapport nicht erfasst - überspringen
                    Case ""
                        ' Diese werden im Rapport nicht erfasst - überspringen
                    Case Else
                        '----------------------------------------------
                        ' Reguläres Projekt: Stunden eintragen
                        '----------------------------------------------
                        
                        ' Prüfen ob Projekt bereits im Rapport vorhanden ist
                        Set MABProj = wsMAB.UsedRange.Resize(, 1).offset(0, 13)
                        ' Kommentar extrahieren wenn Zeilenumbruch vorhanden
                            strComment = ""
                            If InStr(cell.Value, Chr(10)) > 0 Then
                                carr = Split(cell.Value, Chr(10))
                                If UBound(carr) >= 1 Then
                                    strProj = carr(0)
                                    strComment = carr(1)  ' Zweite Zeile als Kommentar
                                End If
                            Else
                                strProj = cell.Value
                            End If
                            
                        Debug.Print "Search ", strProj, "in " & MABProj.Address
                        
                        If Not MABProj.Find(strProj) Is Nothing Then
                            ' Projekt existiert bereits - Zeile finden
                            MABProjrow = MABProj.Find(strProj).row
                            With wsMAB.Cells(MABProjrow, i + 2)
                                .Value = 8.5
                                If Len(strComment) > 0 Then ' wenn ein Kommentar geschrieben werden soll
                                Debug.Print "addComment ", strComment
                                    .AddComment
                                    .comment.Text strComment
                                    '.comment.Visible = True
                                End If
                            End With
                        Else
                            ' Neues Projekt - neue Zeile hinzufügen
                            MABProjrow = wsMAB.Cells(24, 1).End(xlUp).row + 1
                            
                            ' Projektname in Spalte N eintragen
                            wsMAB.Cells(MABProjrow, 14).Value = strProj
                            
                            ' Bemerkung und Kommissionsnummer aus Dictionary auslesen
                            larr = Split(PROJEKTE(strProj), ";")
                            
                            If UBound(larr) > 0 Then
                                wsMAB.Cells(MABProjrow, 1).Value = larr(1) ' Bemerkung
                                wsMAB.Cells(MABProjrow, 2).Value = larr(0) ' Kommissionsnummer
                            Else
                                ' Fallback wenn keine Bemerkung vorhanden
                                wsMAB.Cells(MABProjrow, 1).Value = PROJEKTE(cell.Value)
                            End If
                            
                            ' Stunden für diesen Tag eintragen
                            With wsMAB.Cells(MABProjrow, i + 2)
                                .Value = 8.5
                                If Len(strComment) > 0 Then ' wenn ein Kommentar geschrieben werden soll
                                    .AddComment
                                    .comment.Text strComment
                                    '.comment.Visible = True
                                End If
                            End With
                        End If
                End Select
            Next cell
        End If
        
SkipMAB:
        On Error GoTo ErrorHandler ' Fehlerbehandlung wieder aktivieren
    Next lkey
    
    '==========================================================================
    ' Aufräumen und Abschluss
    '==========================================================================
CleanupAndExit:
    
    ' Leeres erstes Blatt in neuer Arbeitsmappe löschen
    If newWB.Sheets.Count > 1 Then
        Application.DisplayAlerts = False
        newWB.Sheets(1).Delete
        Application.DisplayAlerts = True
    End If
    
    ' Neue Arbeitsmappe neu berechnen
    newWB.Application.Calculate
    
    ' Originaleinstellungen wiederherstellen
    With Application
        .ScreenUpdating = originalScreenUpdating
        .Calculation = originalCalculation
        .DisplayAlerts = originalDisplayAlerts
        .EnableEvents = True
    End With
    
    wsTemplate.Visible = xlSheetHidden
    
    ' Erfolgsmeldung anzeigen
    MsgBox cntRapporte & " Wochenrapporte wurden erfolgreich erstellt!", _
           vbInformation, "Rapporte erstellt"
    
    Exit Sub
    
    '==========================================================================
    ' Fehlerbehandlung
    '==========================================================================
ErrorHandler:
    ' Originaleinstellungen auch bei Fehler wiederherstellen
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    
    ' Fehlermeldung anzeigen
    MsgBox "Fehler " & Err.Number & " ist aufgetreten:" & vbNewLine & vbNewLine & _
           Err.Description & vbNewLine & vbNewLine & _
           "Quelle: " & Err.Source, _
           vbCritical, "Fehler bei der Rapport-Erstellung"
    
    ' Fehler protokollieren (optional)
    Debug.Print "Fehler in WR_Erstellen: " & Err.Number & " - " & Err.Description
    
End Sub



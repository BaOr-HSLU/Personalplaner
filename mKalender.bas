Attribute VB_Name = "mKalender"
'@Folder "Personalplaner"
'@ModuleDescription "Kalender erstellen und formatieren"
Option Explicit

Const anzahlZeilen                     As Long = 50

Sub ErstelleKalenderMitArbeitstagen(ByVal startZelle As Range)
    Dim ws                             As Worksheet
    Dim startDatum                     As Date
    Dim endDatum                       As Date
    Dim aktDatum                       As Date
    Dim spalte                         As Long
    Dim zeile                          As Long
    Dim kwStartSpalte                  As Long
    Dim monatStartSpalte               As Long
    Dim jahrStartSpalte                As Long
    Dim aktuelleKW                     As Long
    Dim aktuellerMonat                 As String
    Dim aktuellesJahr                  As String
    Dim ersterTagKW                    As Date
    Dim letzterTagKW                   As Date
    Dim zeileOffsetTage                As Long
    Dim zeileOffsetKW                  As Long
    Dim zeileOffsetMonat               As Long
    Dim zeileOffsetFerien              As Long
    Dim zeileOffsetFeiertage           As Long
    
    Set ws = ActiveSheet
    
    If startZelle Is Nothing Then
        MsgBox "Keine Startzelle ausgewählt.", vbExclamation
        Exit Sub
    End If
    
    ' Datumseingabe
    startDatum = Application.InputBox("Startdatum eingeben (z.B. 01.01.2025):", "Startdatum", Date, , , , , 1)
    endDatum = Application.InputBox("Enddatum eingeben (z.B. 31.12.2025):", "Enddatum", Date + 30, , , , , 1)
    
    If endDatum < startDatum Then
        MsgBox "Enddatum muss nach dem Startdatum liegen!", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    
    spalte = startZelle.Column
    zeile = startZelle.row
    
    aktDatum = startDatum
    aktuelleKW = WorksheetFunction.WeekNum(aktDatum, 2)
    aktuellerMonat = Format(aktDatum, "MMMM")
    aktuellesJahr = Format(aktDatum, "YYYY")
    
    kwStartSpalte = spalte
    monatStartSpalte = spalte
    jahrStartSpalte = spalte
    
    zeileOffsetTage = -1
    zeileOffsetKW = -2
    zeileOffsetMonat = -3
    zeileOffsetFeiertage = -5
    zeileOffsetFerien = -4
    
    With ws.Range(ws.Cells(zeile + zeileOffsetTage, spalte), ws.Cells(zeile + anzahlZeilen, spalte)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    
    Do While aktDatum <= endDatum
        ' Nur Montag bis Freitag
        If Weekday(aktDatum, vbMonday) <= 5 Then
            ' Datum eintragen
            ws.Cells(zeile, spalte).Value = aktDatum
            ws.Cells(zeile, spalte).NumberFormat = "dd"
            ws.Cells(zeile, spalte).HorizontalAlignment = xlCenter
            ws.Columns(spalte).ColumnWidth = 0.69
            
            With ws.Range(ws.Cells(zeile - 5, spalte), ws.Cells(zeile - 8, spalte))
                .Merge
                .Font.Size = 6
                .Orientation = 90
                .VerticalAlignment = xlBottom
                .HorizontalAlignment = xlCenter
            End With
            
            ' Vertikale Linien > zwischen den Linien
            'With ws.Range(ws.Cells(zeile, spalte), ws.Cells(zeile + anzahlZeilen, spalte)).Borders(xlEdgeLeft)
            '    .LineStyle = xlNone
            '    .Weight = xlThin
            '    .Color = RGB(100, 100, 100)
            'End With
            
            ' KW-Wechsel
            If WorksheetFunction.WeekNum(aktDatum, 2) <> aktuelleKW Then
                ' KW-Zeile
                With ws.Range(ws.Cells(zeile + zeileOffsetKW, kwStartSpalte), ws.Cells(zeile + zeileOffsetKW, spalte - 1))
                    .Merge
                    .Value = CStr(aktuelleKW)
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True
                    .Font.Size = 10
                    With .Borders
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                        .Color = RGB(0, 0, 0)
                    End With
                End With
                
                
                ' Datumsbereich für diese Woche
                ersterTagKW = ws.Cells(zeile, kwStartSpalte).Value
                letzterTagKW = ws.Cells(zeile, spalte - 1).Value
                With ws.Range(ws.Cells(zeile + zeileOffsetTage, kwStartSpalte), ws.Cells(zeile + zeileOffsetTage, spalte - 1))
                    .Merge
                    .Value = Format(ersterTagKW, "dd") & "–" & Format(letzterTagKW, "dd")
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = False
                    .Font.Size = 8
                    With .Borders
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .Color = RGB(0, 0, 0)
                    End With
                End With
                
                ' Linker Rand ab neuer KW
                With ws.Range(ws.Cells(zeile, spalte), ws.Cells(zeile + anzahlZeilen, spalte)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With
                
                ' KW aktualisieren
                aktuelleKW = WorksheetFunction.WeekNum(aktDatum, 2)
                kwStartSpalte = spalte
            End If
            
            ' Monat prüfen
            If Format(aktDatum, "MMMM") <> aktuellerMonat Then
                With ws.Range(ws.Cells(zeile + zeileOffsetMonat, monatStartSpalte), ws.Cells(zeile + zeileOffsetMonat, spalte - 1))
                    .Merge
                    .Value = CStr(aktuellerMonat) & " " & CStr(aktuellesJahr)
                    .NumberFormat = "MMMM YYYY"
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True
                    .Font.Size = 11
                    With .Borders
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                        .Color = RGB(0, 0, 0)
                    End With
                End With
                
                aktuellerMonat = Format(aktDatum, "MMMM")
                monatStartSpalte = spalte
            End If
            
            ' Jahr prüfen
            If Format(aktDatum, "YYYY") <> aktuellesJahr Then
                aktuellesJahr = Format(aktDatum, "YYYY")
            End If
            
            spalte = spalte + 1
        End If
        Application.StatusBar = aktuellesJahr & " / " & aktuellerMonat & " / " & aktuelleKW & " / " & aktDatum
        aktDatum = aktDatum + 1
    Loop
    
    ' Letzte KW abschließen
    With ws.Range(ws.Cells(zeile + zeileOffsetKW, kwStartSpalte), ws.Cells(zeile + zeileOffsetKW, spalte - 1))
        .Merge
        .Value = CStr(aktuelleKW)
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 10
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
    End With
    
    ersterTagKW = ws.Cells(zeile, kwStartSpalte).Value
    letzterTagKW = ws.Cells(zeile, spalte - 1).Value
    With ws.Range(ws.Cells(zeile + zeileOffsetTage, kwStartSpalte), ws.Cells(zeile + zeileOffsetTage, spalte - 1))
        .Merge
        .Value = Format(ersterTagKW, "dd") & "–" & Format(letzterTagKW, "dd")
        .HorizontalAlignment = xlCenter
        .Font.Bold = False
        .Font.Size = 8
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End With
    
    ' Letzten Monatsnamen abschließen
    With ws.Range(ws.Cells(zeile + zeileOffsetMonat, monatStartSpalte), ws.Cells(zeile + zeileOffsetMonat, spalte - 1))
        .Merge
        .Value = CStr(aktuellerMonat) & " " & CStr(aktuellesJahr)
        .NumberFormat = "MMMM YYYY"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 11
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
    End With
    
    ' Rechte Rahmenlinie ganz rechts
    'With ws.Range(ws.Cells(zeile, spalte - 1), ws.Cells(zeile + anzahlZeilen, spalte - 1)).Borders(xlEdgeRight)
    '    .LineStyle = xlContinuous
    '    .Weight = xlMedium
    '    .Color = RGB(0, 0, 0)
    'End With
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    
    On Error Resume Next
    ThisWorkbook.Names("TAGE").Delete
    On Error GoTo 0
    
    ThisWorkbook.Names.Add Name:="TAGE", RefersTo:=ActiveSheet.Range(ActiveSheet.Cells(startZelle.row, startZelle.Column), ActiveSheet.Cells(startZelle.row, spalte - 1))
    
    MsgBox "Kalender mit Arbeitstagen erfolgreich erstellt!", vbInformation
    
    Select Case MsgBox("Sollgen die Feiertage auch eingetragen werden?", vbYesNo, "Feiertage eintragen")
        Case vbYes
            FerienUndFeiertageEintragen
    End Select
    
    BedingteFormatierungMitDropdownsInTabellen
    
    Tabelle1.Activate
    
    Application.StatusBar = False
End Sub

Public Sub FerienUndFeiertageEintragen()
    Dim DatumStart                     As Date
    Dim DatumEnd                       As Date
    Dim FName                          As String
    Dim rngDatum                       As Range
    Dim rngFerien                      As Range
    Dim fRow                           As ListRow
    Dim loFeiertage                    As ListObject
    Dim loFerien                       As ListObject
    Dim cell                           As Range
    Dim colStart                       As Long
    Dim colEnd                         As Long
    Dim TAGE_ROW                       As Long
    
    ' Position der Zeile mit den Datumswerten ("TAGE")
    TAGE_ROW = ActiveSheet.Range("TAGE").Rows(1).row
    
    ' SCHULFERIEN
    Set loFerien = Tabelle1.ListObjects("Ferien")
    
    For Each fRow In loFerien.ListRows
        DatumStart = fRow.Range.Cells(1, 2).Value
        DatumEnd = fRow.Range.Cells(1, 3).Value
        FName = fRow.Range.Cells(1, 1).Value
        
        Application.StatusBar = "Ferien / " & FName & " von " & DatumStart & " bis " & DatumEnd
        
        colStart = 0
        colEnd = 0
        
        ' Durchsuche alle Zellen im Bereich "TAGE"
        For Each cell In ActiveSheet.Range("TAGE").Cells
            If IsDate(cell.Value) Then
                If cell.Value >= DatumStart And cell.Value <= DatumEnd Then
                    ' Bereich unterhalb einfärben
                    'With Tabelle7.Range( _
                    '    Tabelle7.Cells(TAGE_ROW, cell.Column), _
                    '    Tabelle7.Cells(TAGE_ROW + ZEILEN, cell.Column) _
                    ').Interior
                    '    .Pattern = xlSolid
                    '    .ThemeColor = xlThemeColorAccent2
                    '    .TintAndShade = 0.8
                    'End With
                    
                    ' Start- und Endspalten merken
                    If colStart = 0 Then colStart = cell.Column
                    colEnd = cell.Column
                End If
            End If
        Next cell
        
        ' Merge + Ferienname nur, wenn Spalten vorhanden
        If colStart > 0 And colEnd >= colStart Then
            With ActiveSheet.Range( _
                 ActiveSheet.Cells(TAGE_ROW - 4, colStart), _
                 ActiveSheet.Cells(TAGE_ROW - 4, colEnd) _
                 )
                .Merge
                .Value = FName
                .Font.Size = 6
                .HorizontalAlignment = xlCenter
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(0, 0, 0)
                End With
            End With
        End If
    Next fRow
    
    ' FEIERTAGE
    Set loFeiertage = Tabelle1.ListObjects("Feiertage")
    
    For Each fRow In loFeiertage.ListRows
        DatumStart = fRow.Range.Cells(1, 2).Value
        FName = fRow.Range.Cells(1, 1).Value
        Set rngDatum = Nothing
        
        Application.StatusBar = "Feiertag / " & FName & " " & DatumStart
        
        ' Manuelle Suche nach Datum im Bereich "TAGE"
        For Each cell In ActiveSheet.Range("TAGE").Cells
            If IsDate(cell.Value) Then
                If CLng(cell.Value) = CLng(DatumStart) Then
                    Set rngDatum = cell
                    Exit For
                End If
            End If
        Next cell
        
        If Not rngDatum Is Nothing Then
            ' Bereich unterhalb der TAGE-Zeile einfärben (z. B. 6 Zeilen darunter)
            With ActiveSheet.Range( _
                 ActiveSheet.Cells(TAGE_ROW, rngDatum.Column), _
                 ActiveSheet.Cells(TAGE_ROW + anzahlZeilen, rngDatum.Column) _
                 ).Interior
                .Pattern = xlSolid
                .ColorIndex = 33
            End With
            
            ' Feiertagsname oberhalb eintragen (z. B. in Zeile 2)
            With ActiveSheet.Cells(TAGE_ROW - 8, rngDatum.Column)
                .Value = FName
                .Interior.Pattern = xlSolid
                .Interior.ColorIndex = 33
            End With
        Else
Debug.Print "Feiertag NICHT gefunden", FName, DatumStart
        End If
    Next fRow
    
    MsgBox "Feiertage und Schulferien wurden erfolgreich eingetragen.", vbInformation
    Application.StatusBar = False
End Sub

Sub BedingteFormatierungMitDropdownsInTabellen(Optional ByVal Kurzform As Boolean = True, Optional ByVal startcol As Long = 15)
    Dim ws                             As Worksheet
    Dim lo                             As ListObject
    Dim iCol                           As Long
    Dim ZielBereich                    As Range
    Dim tmpRange                       As Range
    Dim formel                         As String
    Dim validationList                 As String
    Dim wert                           As Variant
    Dim ersetzungen                    As Object
    Set ersetzungen = CreateObject("Scripting.Dictionary")
    
    ' Kürzel ? Langform
    ersetzungen.Add "Fx", "Ferien nicht bewilligt"
    ersetzungen.Add "F", "Ferien"
    ersetzungen.Add "U", "Unfall"
    ersetzungen.Add "K", "Krank"
    ersetzungen.Add "WK", "Militär"
    ersetzungen.Add "S", "Schule"
    ersetzungen.Add "ÜK", "Überbetr. Kurs"
    ersetzungen.Add "T", "Teilzeit"
    
    ' Dropdown-Werteliste (je nach Einstellung)
    If Kurzform Then
        validationList = Join(ersetzungen.Keys, ",")
    Else
        validationList = Join(ersetzungen.Items, ",")
    End If
    
    Set ws = ActiveSheet
    Set ZielBereich = Nothing
    
    ' === Schritt 1: Bereiche zusammenstellen (ab Spalte 15 in allen Tabellen) ===
    For Each lo In ws.ListObjects
        For iCol = startcol To lo.Range.Columns.Count
            Set tmpRange = lo.ListColumns(iCol).DataBodyRange
            If Not ZielBereich Is Nothing Then
                Set ZielBereich = Union(ZielBereich, tmpRange)
            Else
                Set ZielBereich = tmpRange
            End If
        Next iCol
    Next lo
    
    If ZielBereich Is Nothing Then
        MsgBox "Keine gültigen Zellen ab Spalte 15 gefunden.", vbExclamation
        Exit Sub
    End If
    
    ' === Schritt 2: Bedingte Formatierungen löschen und neu hinzufügen ===
    ZielBereich.FormatConditions.Delete
    
    For Each wert In ersetzungen.Keys
        ' Formel richtet sich nach ausgewählter Form
        If Kurzform Then
            formel = "=" & ZielBereich.Cells(1, 1).Address(False, False) & "=""" & wert & """"
        Else
            formel = "=" & ZielBereich.Cells(1, 1).Address(False, False) & "=""" & ersetzungen(wert) & """"
        End If
        
        With ZielBereich.FormatConditions.Add(Type:=xlExpression, Formula1:=formel)
            .StopIfTrue = False
            Select Case wert
                Case "Fx": .Interior.Color = RGB(255, 0, 0)
                Case "F":  .Interior.Color = RGB(0, 176, 240)
                Case "U":  .Interior.Color = RGB(255, 192, 0)
                Case "K":  .Interior.Color = RGB(255, 192, 0)
                Case "WK": .Interior.Color = RGB(0, 176, 80)
                Case "S":  .Interior.Color = RGB(224, 237, 201)
                Case "ÜK": .Interior.Color = RGB(224, 237, 201)
                Case "T":  .Interior.Color = RGB(191, 191, 191)
            End Select
        End With
    Next wert
    
    ' === Schritt 3: Dropdown mit freier Eingabe ===
    With ZielBereich.Validation
        .Delete
    End With
    
    'ZielBereich.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Operator:=xlBetween, Formula1:=validationList
    'ZielBereich.Validation.IgnoreBlank = True
    'ZielBereich.Validation.InCellDropdown = True
    'ZielBereich.Validation.ShowError = False
    'ZielBereich.Validation.ShowInput = True
    'ZielBereich.Validation.InputTitle = "Eingabe"
    'ZielBereich.Validation.InputMessage = "Wähle aus der Liste oder gib frei ein."
    
    'MsgBox "Formatierung und Dropdowns wurden mit " & IIf(Kurzform, "KURZ", "LANG") & "form angewendet.", vbInformation
End Sub



Attribute VB_Name = "mKWBlatt"
'@Folder "Personalplaner"
'@ModuleDescription "Kalenderwoche erstellen und formatieren."
Option Explicit

Sub test()
    NeuesKWBlattErstellen ActiveSheet.Range("GZ8")
End Sub

Public Sub NeuesKWBlattErstellen(Target As Range)
    Application.ScreenUpdating = False
    Application.StatusBar = "Wochenplan erstellen ..."
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim KW                             As Long
    On Error Resume Next
    KW = CLng(Target.offset(0, 0).Value)
    On Error GoTo 0
    If KW = 0 Then
        MsgBox "Keine gültige Kalenderwoche ausgewählt!", vbExclamation
        Exit Sub
    End If
    
    Dim KWStart                        As Date
    KWStart = Target.offset(2, 0).Value
    Dim KWEnd                          As Date
    KWEnd = Target.offset(2, 0).offset(0, 4).Value
    
    Dim zielName                       As String
    zielName = "KW" & KW & " " & Format(KWStart, "YYYY")
    
    Dim ws                             As Worksheet
    ' Altes Blatt aktivieren, wenn es existiert
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = zielName Then
            ws.Visible = xlSheetVisible
            ws.Activate
            Exit Sub
        End If
    Next ws
    
    ' Vorlage sichtbar machen, falls veryhidden
    With Tabelle7
        .copying = True
        .Visible = xlSheetVisible
        .Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set ws = ActiveSheet
        On Error Resume Next
        ws.Name = zielName
        On Error GoTo 0
        .Visible = xlSheetHidden
        .copying = False
    End With
    
    ws.Range("A3:A4").Value = "KW" & KW
    ws.Range("E4").Value = KWStart
    ws.Range("F4").Value = KWEnd
    ws.Range("J3").Value = Now()
    ' Vorlage wieder veryhidden setzen (optional)
    'Tabelle7.Visible = xlSheetVeryHidden
    
    Application.StatusBar = "Das Blatt '" & zielName & "' wurde neu erstellt."
    
    ' Hier kannst du jetzt deine Logik einfügen,
    ' um die Daten für die KW zu kopieren (wie vorher besprochen)
    
    Dim startcol                       As Long
    Dim endcol                         As Long
    startcol = Target.Column
    endcol = startcol + 4
    Dim i                              As Long, j As Long
    Dim lo                             As ListObject
    Dim loKw                           As ListObject
    Set loKw = ws.Range("A7").ListObject
    Dim lrow                           As ListRow
    Dim lrowkw                         As ListRow
    Application.ScreenUpdating = False
    For Each lo In Tabelle3.ListObjects
        For Each lrow In lo.ListRows
            If lrow.Range(1, 7).Value = vbNullString Then GoTo SKIPLROW
            Set lrowkw = loKw.ListRows.Add
            lrowkw.Range(1, 1).Value = lrow.Range(1, 6).Value ' Nummer
            lrowkw.Range(1, 2).Value = lrow.Range(1, 7).Value & vbNewLine & lrow.Range(1, 9).Value & vbNewLine & lrow.Range(1, 13).Value ' Name & Telefonnummer & E-Mail
            Debug.Print "Databody", lo.DataBodyRange.Address, "Lrow", lrow.Range(1, 7).Address, "lrowKW", lrowkw.Range(1, 1).Address
            Debug.Print , lrow.Range(1, 7).Value, lrow.Range(1, 9).Value, lrow.Range(1, 13).Value
            ErsteZeileImBereichFett lrowkw.Range(1, 2)
            lrowkw.Range(1, 3).Value = lrow.Range(1, 8).Value ' Funktion
            lrowkw.Range(1, 4).Value = lrow.Range(1, 10).Value ' Team
            j = 5
            For i = startcol To endcol
                lrowkw.Range(1, j).Value = lrow.Range(1, i).Value
                j = j + 1
            Next i
SKIPLROW:
        Next lrow
    Next lo
    
    Dim ersetzungen                    As Object
    Set ersetzungen = CreateObject("Scripting.Dictionary")
    
    ' Nur definierte Kürzel werden ersetzt
    ersetzungen.Add "Fx", "Ferien nicht bewilligt"
    ersetzungen.Add "F", "Ferien"
    ersetzungen.Add "U", "Unfall"
    ersetzungen.Add "K", "Krank"
    ersetzungen.Add "WK", "Militär"
    ersetzungen.Add "S", "Schule"
    ersetzungen.Add "ÜK", "Überbetr. Kurs"
    ersetzungen.Add "T", "Teilzeit"
    
    Dim cell                           As Range
    For Each cell In loKw.DataBodyRange()        ' oder dein konkreter Bereich
        If ersetzungen.Exists(cell.Value) Then
            cell.Value = ersetzungen(cell.Value)
        End If
        ' Sonstige Werte (z. B. "KiSpi") bleiben erhalten
    Next cell
    
    BedingteFormatierungMitDropdownsInTabellen False, 5
    
    InitListBox ws, "ListBoxFunktion", "Funktion"
    InitListBox ws, "ListBoxTeam", "Team"
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculate
    Application.StatusBar = "DONE!"
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' ===========================================================================
' Hilfsroutine: befüllt eine ActiveX-ListBox aus einer Tabellenspalte
' ===========================================================================
Public Sub InitListBox(ws As Worksheet, listBoxName As String, columnName As String)
    Application.StatusBar = "Initialisiere " & listBoxName
    Dim lo                             As ListObject
    Set lo = ws.ListObjects(1)                   ' erste Tabelle im neuen Blatt
    
    Dim lb                             As MSForms.Listbox
    On Error Resume Next
    Set lb = ws.OLEObjects(listBoxName).Object
    On Error GoTo 0
    
    If lb Is Nothing Then
        MsgBox "ListBox '" & listBoxName & "' nicht gefunden auf Blatt " & ws.Name, vbExclamation
        Exit Sub
    End If
    
    Dim rng                            As Range, cell As Range
    On Error Resume Next
    Set rng = lo.ListColumns(columnName).DataBodyRange
    On Error GoTo 0
    
    If rng Is Nothing Then
        lb.Clear
        'MsgBox "Spalte '" & columnName & "' nicht gefunden in Tabelle '" & lo.Name & "'.", vbExclamation
        Exit Sub
    End If
    
    ' ListBox zurücksetzen
    lb.Clear
    
    Dim dict                           As New Dictionary
    Set dict = SammleEindeutigeWerteSchnellRng(rng, OnlyFirstLine:=True)
    
    Dim Key                            As Variant
    For Each Key In dict.Keys
        lb.AddItem (CStr(Key))
    Next Key
    
End Sub

Function AnfangsspalteVorherigeKW(Optional ByVal AktuelleSpalte As Long = 0) As Long
    Dim ws                             As Worksheet
    Set ws = ActiveSheet
    
    Dim col                            As Long
    Dim aktuelleKW                     As Long
    Dim aktuelleDatum                  As Date
    Dim vorherigeKW                    As Long
    Dim vorherigesJahr                 As Long
    Dim zielSpalte                     As Long
    Dim gefunden                       As Boolean
    
    ' Datum in Zeile 10 (Werktage) lesen
    aktuelleDatum = Format(Now(), "DD.MM.YYYY")
    If Not IsDate(aktuelleDatum) Then
        MsgBox "In Spalte " & AktuelleSpalte & " steht kein gültiges Datum in Zeile 10.", vbExclamation
        Exit Function
    End If
    
    aktuelleKW = WorksheetFunction.IsoWeekNum(aktuelleDatum)
    vorherigeKW = aktuelleKW - 1
    vorherigesJahr = Year(aktuelleDatum)
    
    ' Wenn wir in KW 1 sind ? Vorjahr behandeln
    If vorherigeKW < 1 Then
        vorherigesJahr = vorherigesJahr - 1
        vorherigeKW = WorksheetFunction.IsoWeekNum(DateSerial(vorherigesJahr, 12, 31)) ' letzte KW des Vorjahres
    End If
    
    ' Suche in Zeile 8 die Zelle mit der vorherigen KW (Merged-Zellen beachten)
    For col = 1 To ws.Cells(8, ws.Columns.Count).End(xlToLeft).Column
        Dim zelleKW                    As Range
        Set zelleKW = ws.Cells(8, col)
        
        ' Nur die linke obere Zelle eines Merge-Bereichs prüfen
        If zelleKW.MergeCells And zelleKW.Address = zelleKW.MergeArea.Cells(1, 1).Address Then
            Dim kwWert                 As Variant
            kwWert = zelleKW.Value
            If IsNumeric(kwWert) And Not IsEmpty(kwWert) Then
                Dim testDatum          As Date
                testDatum = ws.Cells(10, col).Value
                If IsDate(testDatum) Then
                    Dim testKW         As Long
                    Dim testJahr       As Long
                    testKW = WorksheetFunction.IsoWeekNum(testDatum)
                    testJahr = Year(testDatum)
                    
                    If testKW = vorherigeKW And testJahr = vorherigesJahr Then
                        AnfangsspalteVorherigeKW = col
                        gefunden = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next col
    
    If Not gefunden Then
        MsgBox "Die vorherige Kalenderwoche (" & vorherigeKW & "/" & vorherigesJahr & ") wurde nicht gefunden.", vbExclamation
    End If
End Function


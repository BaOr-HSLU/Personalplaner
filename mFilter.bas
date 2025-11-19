Attribute VB_Name = "mFilter"
Option Explicit

' ===========================================================================
' Ereignisse der ActiveX-ListBoxen zum Filtern der Tabelle
' ===========================================================================
' Annahmen:
'   - Tabelle ab Zeile 7 (wie bei dir bisher)
'   - Spalten heißen "Mitarbeiter", "Funktion", "Team"
'   - ListBox-Namen: ListBoxFunktion, ListBoxTeam
' ===========================================================================

' ===========================================================================
' Hilfsroutine: filtert die Tabelle anhand der Auswahl in einer ListBox
' ===========================================================================
Public Sub ApplyTableFilter(ws As Worksheet, listBoxName As String, columnName As String)
    Dim lo                             As ListObject
    Set lo = ws.ListObjects(1)                   ' erste Tabelle im Blatt
    
    Dim lb                             As MSForms.Listbox
    Set lb = ws.OLEObjects(listBoxName).Object
    
    Dim selItems                       As Collection
    Set selItems = New Collection
    
    Dim i                              As Long
    For i = 0 To lb.ListCount - 1
        If columnName = "Mitarbeiter" Then
            If lb.Selected(i) Then selItems.Add lb.List(i) & "*"
        Else
            If lb.Selected(i) Then selItems.Add lb.List(i)
        End If
    Next i
    
    ' Filter zurücksetzen, wenn nichts gewählt
    If selItems.Count = 0 Then
        On Error Resume Next
        lo.Range.AutoFilter Field:=lo.ListColumns(columnName).Index
        On Error GoTo 0
        Application.Calculate
        Exit Sub
    End If
    
    ' Array mit Auswahlwerten bauen
    Dim arr()                          As Variant
    ReDim arr(1 To selItems.Count)
    For i = 1 To selItems.Count
        arr(i) = selItems(i)
    Next i
    
    ' Filter anwenden
    lo.Range.AutoFilter Field:=lo.ListColumns(columnName).Index, _
                        Criteria1:=arr, _
                        Operator:=xlFilterValues
    
    Application.Calculate
End Sub



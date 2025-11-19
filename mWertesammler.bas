Attribute VB_Name = "mWertesammler"
'@Folder "Personalplaner"
'@ModuleDescription "Globale Variabeln"
Public Function SammleEindeutigeWerteSchnell(colStart As Long, Optional onlyFirstRow As Boolean = False) As Dictionary
    Dim ws                             As Worksheet
    Dim lo                             As ListObject
    Dim loArray                        As Variant
    Dim dict                           As Object
    Dim i                              As Long, j As Long
    Dim tempWert                       As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Alle Tabellen durchgehen
    Set ws = ActiveSheet
    For Each lo In ws.ListObjects
        If Not lo.DataBodyRange Is Nothing Then
            loArray = lo.DataBodyRange.Value
            For i = 1 To UBound(loArray, 1)
                For j = colStart To UBound(loArray, 2)
                    tempWert = Trim(CStr(loArray(i, j)))
                    If Len(tempWert) > 0 Then
                        If onlyFirstRow Then
                            If InStr(tempWert, Chr(10)) > 0 Then
                                ' Enthält Zeilenumbruch - nimm erste Zeile
                                If Not dict.Exists(Split(tempWert, Chr(10))(0)) Then dict.Add Split(tempWert, Chr(10))(0), vbNullString
                            Else
                                ' Kein Zeilenumbruch - nimm kompletten Wert
                                If Not dict.Exists(tempWert) Then dict.Add tempWert, vbNullString
                            End If
                        Else
                            If Not dict.Exists(tempWert) Then dict.Add tempWert, vbNullString
                        End If
                    End If
                Next j
            Next i
        End If
    Next lo
    
    Set SammleEindeutigeWerteSchnell = SortDictionaryAlphabetical(dict)
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
End Function

Public Function SammleEindeutigeWerteSchnellRng(rng As Range, Optional ByVal includeHidden As Boolean = True, Optional OnlyFirstLine As Boolean = False) As Dictionary
    Dim ws                             As Worksheet
    Dim loArray                        As Variant
    Dim dict                           As Object
    Dim i                              As Long, j As Long
    Dim tempWert                       As String
    Dim rngRow                         As Range
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' --- Fall 1: Verborgene Zeilen werden berücksichtigt ---
    If includeHidden = True Then
        loArray = rng.Value
        For i = 1 To UBound(loArray, 1)
            For j = 1 To UBound(loArray, 2)
                tempWert = Trim$(CStr(loArray(i, j)))
                If Len(tempWert) > 0 Then
                    If OnlyFirstLine Then
                        If InStr(tempWert, Chr(10)) > 0 Then
                            ' Enthält Zeilenumbruch - nimm erste Zeile
                            If Not dict.Exists(Split(tempWert, Chr(10))(0)) Then dict.Add Split(tempWert, Chr(10))(0), vbNullString
                        Else
                            ' Kein Zeilenumbruch - nimm kompletten Wert
                            If Not dict.Exists(tempWert) Then dict.Add tempWert, vbNullString
                        End If
                    Else
                        If Not dict.Exists(tempWert) Then dict.Add tempWert, vbNullString
                    End If

                End If
            Next j
        Next i
        
        ' --- Fall 2: Verborgene Zeilen werden ignoriert ---
    Else
        For Each rngRow In rng.Rows
            If rngRow.EntireRow.Hidden = False Then
                For j = 1 To rngRow.Columns.Count
                    tempWert = Trim$(CStr(rngRow.Cells(1, j).Value))
                    If Len(tempWert) > 0 Then
                        If OnlyFirstLine Then
                            If InStr(tempWert, Chr(10)) > 0 Then
                                ' Enthält Zeilenumbruch - nimm erste Zeile
                                If Not dict.Exists(Split(tempWert, Chr(10))(0)) Then dict.Add Split(tempWert, Chr(10))(0), vbNullString
                            Else
                                ' Kein Zeilenumbruch - nimm kompletten Wert
                                If Not dict.Exists(tempWert) Then dict.Add tempWert, vbNullString
                            End If
                        Else
                            If Not dict.Exists(tempWert) Then dict.Add tempWert, vbNullString
                        End If
                    End If
                Next j
            End If
        Next rngRow
    End If
    
    ' --- Ergebnis sortieren ---
    Set SammleEindeutigeWerteSchnellRng = SortDictionaryAlphabetical(dict)
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
End Function


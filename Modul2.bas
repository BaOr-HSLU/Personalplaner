Attribute VB_Name = "Modul2"
'@Folder("TOOLS")
'@ModuleDescription "Hilfsfunktionen zum Sortieren von Dictionary-Keys alphabetisch."
Option Explicit

' ============================================================================
' @Description Sortiert die Keys eines Dictionary alphabetisch (aufsteigend).
' @Param dict Das unsortierte Dictionary-Objekt.
' @Return Ein neues Dictionary, dessen Einträge alphabetisch nach Keys sortiert sind.
' ============================================================================
Public Function SortDictionaryAlphabetical(ByVal dict As Object) As Object
    On Error GoTo ErrHandler
    
    Dim arrKeys()                      As String
    Dim i                              As Long, j As Long
    Dim temp                           As String
    Dim sortedDict                     As Object
    
    ' Dictionary prüfen
    If dict Is Nothing Then
        Set SortDictionaryAlphabetical = Nothing
        Exit Function
    End If
    
    If dict.Count = 0 Then
        Set SortDictionaryAlphabetical = dict
        Exit Function
    End If
    
    ' Keys in Array schreiben
    ReDim arrKeys(0 To dict.Count - 1)
    i = 0
    Dim k                              As Variant
    For Each k In dict.Keys
        arrKeys(i) = CStr(k)
        i = i + 1
    Next k
    
    ' Alphabetisch sortieren (BubbleSort für Einfachheit)
    For i = LBound(arrKeys) To UBound(arrKeys) - 1
        For j = i + 1 To UBound(arrKeys)
            If StrComp(arrKeys(i), arrKeys(j), vbTextCompare) > 0 Then
                temp = arrKeys(i)
                arrKeys(i) = arrKeys(j)
                arrKeys(j) = temp
            End If
        Next j
    Next i
    
    ' Neues Dictionary mit sortierten Keys aufbauen
    Set sortedDict = CreateObject("Scripting.Dictionary")
    For i = LBound(arrKeys) To UBound(arrKeys)
        sortedDict.Add arrKeys(i), dict(arrKeys(i))
    Next i
    
    Set SortDictionaryAlphabetical = sortedDict
    Exit Function
    
ErrHandler:
    Set SortDictionaryAlphabetical = Nothing
End Function


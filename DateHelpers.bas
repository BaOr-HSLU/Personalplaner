Attribute VB_Name = "DateHelpers"
'@Folder("Services.Utilities")
'@ModuleDescription("Date and calendar week helper functions")
Option Explicit

'@Description("Finds column containing a specific date in the header row. Handles date formats, text dates, and dates with time")
'@Param targetSheet The worksheet to search in
'@Param headerRowNumber The row number containing date headers
'@Param searchDate The date to find (time portion is ignored)
'@Param firstColumnToSearch Optional first column to search (default 1)
'@Param lastColumnToSearch Optional last column to search (0 = auto-detect)
'@Returns Column number (1-based) or 0 if not found
Public Function FindDateColumn( _
        ByVal targetSheet As Worksheet, _
        ByVal headerRowNumber As Long, _
        ByVal searchDate As Date, _
        Optional ByVal firstColumnToSearch As Long = 1, _
        Optional ByVal lastColumnToSearch As Long = 0) As Long
    '@Ignore EmptyStringLiteral
    On Error GoTo CleanFail

    Dim headerRange As Range
    Dim matchResult As Variant
    Dim searchDateSerial As Double
    Dim currentCell As Range
    Dim autoDetectedLastColumn As Long
    Dim textDateFormat1 As String
    Dim textDateFormat2 As String

    '--- Determine search boundaries
    If lastColumnToSearch = 0 Then
        autoDetectedLastColumn = GetUsedLastColumn(targetSheet)
        If autoDetectedLastColumn = 0 Then GoTo CleanFail
        lastColumnToSearch = autoDetectedLastColumn
    End If

    If firstColumnToSearch < 1 Then firstColumnToSearch = 1
    If lastColumnToSearch > targetSheet.Columns.Count Then lastColumnToSearch = targetSheet.Columns.Count
    If lastColumnToSearch < firstColumnToSearch Then GoTo CleanFail

    '--- Define header range
    Set headerRange = targetSheet.Range(targetSheet.Cells(headerRowNumber, firstColumnToSearch), _
                                        targetSheet.Cells(headerRowNumber, lastColumnToSearch))

    '--- Prepare search values (ignore time portion)
    searchDateSerial = Int(CDbl(searchDate))
    textDateFormat1 = Format$(searchDate, "dd.mm.yyyy")
    textDateFormat2 = Format$(searchDate, "d.m.yyyy")

    '--- Strategy 1: Direct MATCH on numeric date serial
    matchResult = Application.Match(searchDateSerial, headerRange, 0)
    If Not IsError(matchResult) Then
        FindDateColumn = firstColumnToSearch + CLng(matchResult) - 1
        Exit Function
    End If

    '--- Strategy 2: MATCH on text representation
    matchResult = Application.Match(textDateFormat1, headerRange, 0)
    If Not IsError(matchResult) Then
        FindDateColumn = firstColumnToSearch + CLng(matchResult) - 1
        Exit Function
    End If

    matchResult = Application.Match(textDateFormat2, headerRange, 0)
    If Not IsError(matchResult) Then
        FindDateColumn = firstColumnToSearch + CLng(matchResult) - 1
        Exit Function
    End If

    '--- Strategy 3: Manual loop (handles dates with time, convertible text)
    For Each currentCell In headerRange.Cells
        If Not IsError(currentCell.Value2) Then
            If IsDate(currentCell.value) Then
                If Int(CDbl(CDate(currentCell.value))) = searchDateSerial Then
                    FindDateColumn = currentCell.Column
                    Exit Function
                End If
            Else
                ' Try to parse text as date
                If LenB(currentCell.Value2) > 0 Then
                    If IsDate(CStr(currentCell.Value2)) Then
                        If Int(CDbl(CDate(CStr(currentCell.Value2)))) = searchDateSerial Then
                            FindDateColumn = currentCell.Column
                            Exit Function
                        End If
                    Else
                        ' Direct text comparison
                        If CStr(currentCell.Value2) = textDateFormat1 Or CStr(currentCell.Value2) = textDateFormat2 Then
                            FindDateColumn = currentCell.Column
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next currentCell

    '--- Not found
    FindDateColumn = 0
    Exit Function

CleanFail:
    FindDateColumn = 0
End Function

'@Description("Gets the last used column in a worksheet")
'@Param targetSheet The worksheet to check
'@Returns Last used column number or 0 if sheet is empty
Private Function GetUsedLastColumn(ByVal targetSheet As Worksheet) As Long
    '@Ignore EmptyStringLiteral
    On Error GoTo Fail
    GetUsedLastColumn = targetSheet.Columns.Count
    Exit Function
Fail:
    GetUsedLastColumn = 0
End Function

'@Description("Formats multi-line cell - makes first line bold")
'@Param targetRange The range to format
Public Sub FormatFirstLineBold(ByVal targetRange As Range)
    On Error GoTo ErrorHandler

    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim currentCell As Range
    Dim cellText As String
    Dim lineBreakPosition As Long
    Dim firstLineLength As Long

    '--- Save and set application state
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- Validate input
    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 513, "FormatFirstLineBold", "Target range is invalid (Nothing)."
    End If

    '--- Process each cell
    For Each currentCell In targetRange.Cells
        If Not currentCell Is Nothing Then
            If Len(currentCell.Value2) > 0 Then
                cellText = CStr(currentCell.Value2)

                '--- Find line break position
                lineBreakPosition = InStr(1, cellText, vbNewLine, vbBinaryCompare)

                If lineBreakPosition > 0 Then
                    firstLineLength = lineBreakPosition - 1
                Else
                    firstLineLength = Len(cellText)
                End If

                If firstLineLength > 0 Then
                    '--- Reset all to normal, then make first line bold
                    currentCell.Characters().Font.Bold = False
                    currentCell.Characters(Start:=1, Length:=firstLineLength).Font.Bold = True
                End If
            End If
        End If
    Next currentCell

SafeExit:
    '--- Restore application state
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

ErrorHandler:
    MsgBox "Fehler in 'FormatFirstLineBold': " & Err.Number & " - " & Err.Description, _
           vbExclamation, "Formatierungsfehler"
    Resume SafeExit
End Sub

'@Description("Sorts a Dictionary alphabetically by keys")
'@Param sourceDict The dictionary to sort
'@Returns New sorted dictionary
Public Function SortDictionaryAlphabetical(ByVal sourceDict As Dictionary) As Dictionary
    '@Ignore VariableNotUsed
    Dim sortedDict As Dictionary
    Set sortedDict = New Dictionary

    Dim keysArray() As Variant
    ReDim keysArray(0 To sourceDict.Count - 1)

    Dim i As Long
    Dim dictKey As Variant
    For Each dictKey In sourceDict.Keys
        keysArray(i) = dictKey
        i = i + 1
    Next dictKey

    '--- Bubble sort (sufficient for reasonable sizes)
    Dim j As Long
    Dim temp As Variant
    For i = 0 To UBound(keysArray) - 1
        For j = i + 1 To UBound(keysArray)
            If keysArray(i) > keysArray(j) Then
                temp = keysArray(i)
                keysArray(i) = keysArray(j)
                keysArray(j) = temp
            End If
        Next j
    Next i

    '--- Build sorted dictionary
    For i = 0 To UBound(keysArray)
        sortedDict.Add keysArray(i), sourceDict(keysArray(i))
    Next i

    Set SortDictionaryAlphabetical = sortedDict
End Function

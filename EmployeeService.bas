Attribute VB_Name = "EmployeeService"
'@Folder("Services.Employee")
'@ModuleDescription("Service for collecting and managing employee data. Optimized for 200+ employees over 5 years")
Option Explicit

'@Description("Collects unique values from a range using high-performance array processing")
'@Param targetRange The range to scan for unique values
'@Param includeHidden Whether to include hidden rows (default True)
'@Param extractFirstLineOnly Whether to extract only the first line before vbNewLine (default False)
'@Returns Dictionary with unique values as keys
Public Function GetUniqueValuesFromRange( _
        ByVal targetRange As Range, _
        Optional ByVal includeHidden As Boolean = True, _
        Optional ByVal extractFirstLineOnly As Boolean = False) As Dictionary
    '@Ignore EmptyStringLiteral
    Dim uniqueDict As Dictionary
    Set uniqueDict = New Dictionary

    Dim cellValues As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim currentValue As String
    Dim currentRow As Range

    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '--- PERFORMANCE: Process hidden rows with array (fast)
    If includeHidden = True Then
        cellValues = targetRange.value
        For rowIndex = 1 To UBound(cellValues, 1)
            For columnIndex = 1 To UBound(cellValues, 2)
                currentValue = Trim$(CStr(cellValues(rowIndex, columnIndex)))
                If Len(currentValue) > 0 Then
                    If extractFirstLineOnly Then
                        currentValue = ExtractFirstLine(currentValue)
                    End If
                    If Not uniqueDict.Exists(currentValue) Then
                        uniqueDict.Add currentValue, vbNullString
                    End If
                End If
            Next columnIndex
        Next rowIndex

    '--- Process only visible rows (slower, row-by-row)
    Else
        For Each currentRow In targetRange.Rows
            If currentRow.EntireRow.Hidden = False Then
                For columnIndex = 1 To currentRow.Columns.Count
                    currentValue = Trim$(CStr(currentRow.Cells(1, columnIndex).value))
                    If Len(currentValue) > 0 Then
                        If extractFirstLineOnly Then
                            currentValue = ExtractFirstLine(currentValue)
                        End If
                        If Not uniqueDict.Exists(currentValue) Then
                            uniqueDict.Add currentValue, vbNullString
                        End If
                    End If
                Next columnIndex
            End If
        Next currentRow
    End If

    '--- Sort results alphabetically
    Set GetUniqueValuesFromRange = DateHelpers.SortDictionaryAlphabetical(uniqueDict)

    Application.ScreenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
End Function

'@Description("Collects unique values from all ListObjects in the active sheet (optimized for large datasets)")
'@Param startColumnIndex First column index within ListObjects to process
'@Param extractFirstLineOnly Whether to extract only first line before vbNewLine
'@Returns Dictionary with unique values
Public Function GetUniqueValuesFromListObjects( _
        ByVal startColumnIndex As Long, _
        Optional ByVal extractFirstLineOnly As Boolean = False) As Dictionary
    '@Ignore EmptyStringLiteral
    Dim uniqueDict As Dictionary
    Set uniqueDict = New Dictionary

    Dim targetSheet As Worksheet
    Dim listObj As ListObject
    Dim tableData As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim currentValue As String

    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set targetSheet = ActiveSheet

    '--- PERFORMANCE: Read entire table into array, process in memory
    For Each listObj In targetSheet.ListObjects
        If Not listObj.DataBodyRange Is Nothing Then
            tableData = listObj.DataBodyRange.value

            For rowIndex = 1 To UBound(tableData, 1)
                For columnIndex = startColumnIndex To UBound(tableData, 2)
                    currentValue = Trim$(CStr(tableData(rowIndex, columnIndex)))
                    If Len(currentValue) > 0 Then
                        If extractFirstLineOnly Then
                            currentValue = ExtractFirstLine(currentValue)
                        End If
                        If Not uniqueDict.Exists(currentValue) Then
                            uniqueDict.Add currentValue, vbNullString
                        End If
                    End If
                Next columnIndex
            Next rowIndex
        End If
    Next listObj

    Set GetUniqueValuesFromListObjects = DateHelpers.SortDictionaryAlphabetical(uniqueDict)

    Application.ScreenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
End Function

'@Description("Extracts first line from multi-line text (before Chr(10) or vbNewLine)")
'@Param fullText The full text potentially containing line breaks
'@Returns First line only
Private Function ExtractFirstLine(ByVal fullText As String) As String
    If InStr(fullText, Chr(10)) > 0 Then
        ExtractFirstLine = Split(fullText, Chr(10))(0)
    Else
        ExtractFirstLine = fullText
    End If
End Function

'@Description("Parses an Employee object from a multi-line cell in a ListRow")
'@Param employeeRow The ListRow containing employee data
'@Param nameColumnIndex Column index for name (multi-line: Name, Phone, Email)
'@Param functionColumnIndex Column index for job function
'@Param teamColumnIndex Column index for team
'@Param skipColumnIndex Column index for skip flag
'@Returns Employee object
Public Function ParseEmployeeFromListRow( _
        ByVal employeeRow As ListRow, _
        ByVal nameColumnIndex As Long, _
        Optional ByVal functionColumnIndex As Long = 0, _
        Optional ByVal teamColumnIndex As Long = 0, _
        Optional ByVal skipColumnIndex As Long = 0) As Employee

    Dim emp As Employee
    Set emp = New Employee

    '--- Parse multi-line name cell (Name, Phone, Email)
    Dim nameCell As String
    nameCell = employeeRow.Range(1, nameColumnIndex).value
    emp.ParseFromCellValue nameCell

    '--- Optional: Job function
    If functionColumnIndex > 0 Then
        emp.JobFunction = employeeRow.Range(1, functionColumnIndex).value
    End If

    '--- Optional: Team
    If teamColumnIndex > 0 Then
        emp.TeamName = employeeRow.Range(1, teamColumnIndex).value
    End If

    '--- Optional: Skip flag
    If skipColumnIndex > 0 Then
        emp.IsSkipped = employeeRow.Range(1, skipColumnIndex).value
    End If

    Set ParseEmployeeFromListRow = emp
End Function

'@Description("Checks if a cell contains non-empty, non-error value")
'@Param targetCell The cell to check
'@Returns True if cell has usable content
Public Function IsCellNotEmpty(ByVal targetCell As Range) As Boolean
    On Error Resume Next
    IsCellNotEmpty = (Len(Trim$(SafeStringValue(targetCell.Value2))) > 0)
End Function

'@Description("Safely converts a variant to string, handling errors and null values")
'@Param cellValue The cell value to convert
'@Returns String value or empty string
Private Function SafeStringValue(ByVal cellValue As Variant) As String
    On Error Resume Next
    If IsError(cellValue) Or IsNull(cellValue) Or IsEmpty(cellValue) Then
        SafeStringValue = vbNullString
    Else
        SafeStringValue = CStr(cellValue)
    End If
End Function

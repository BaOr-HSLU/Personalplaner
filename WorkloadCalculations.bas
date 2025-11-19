Attribute VB_Name = "WorkloadCalculations"
'@Folder("Services.Formulas")
'@ModuleDescription("User-Defined Functions (UDFs) for workload calculations. Performance-optimized for 200+ employees over 5 years")
Option Explicit

'@Description("Retrieves value from last data row in column matching a date header, with optional column offset")
'@Param targetDate Date to find in header row (time portion ignored)
'@Param columnOffset Offset from date column (0 = exact match)
'@Param headerRowNumber Row number of headers (default 10)
'@Param dataStartRowNumber First data row below headers (default 15)
'@Param anchorColumnNumber Column to find last row (default 1 = column A)
'@Returns Cell value or Excel error (#N/A, #VALUE!, #REF!)
Public Function GetWorkloadByDate( _
        ByVal targetDate As Date, _
        Optional ByVal columnOffset As Long = 0, _
        Optional ByVal headerRowNumber As Long = 10, _
        Optional ByVal dataStartRowNumber As Long = 15, _
        Optional ByVal anchorColumnNumber As Long = 1) As Double

    On Error GoTo ErrorHandler

    Dim targetSheet As Worksheet
    Dim lastDataRow As Long
    Dim dateColumn As Long
    Dim lastUsedColumn As Long

    '--- PERFORMANCE: Use CodeName for speed
    Set targetSheet = Tabelle3

    '--- Validate parameters
    If headerRowNumber < 1 Or dataStartRowNumber < 1 Or dataStartRowNumber <= headerRowNumber Then
        GetWorkloadByDate = CVErr(xlErrValue)
        Exit Function
    End If

    If anchorColumnNumber < 1 Or anchorColumnNumber > targetSheet.Columns.Count Then
        GetWorkloadByDate = CVErr(xlErrRef)
        Exit Function
    End If

    '--- Find last data row in anchor column
    lastDataRow = targetSheet.Cells(targetSheet.Rows.Count, anchorColumnNumber).End(xlUp).Row
    If lastDataRow < dataStartRowNumber Then
        GetWorkloadByDate = CVErr(xlErrNA) 'No data rows
        Exit Function
    End If

    '--- Limit header search to used columns
    lastUsedColumn = GetSheetLastUsedColumn(targetSheet)
    If lastUsedColumn = 0 Then
        GetWorkloadByDate = CVErr(xlErrNA)
        Exit Function
    End If

    '--- Find date column
    dateColumn = DateHelpers.FindDateColumn(targetSheet, headerRowNumber, targetDate, 1, lastUsedColumn)
    If dateColumn = 0 Then
        GetWorkloadByDate = CVErr(xlErrNA) 'Date not found
        Exit Function
    End If

    '--- Apply offset and validate
    dateColumn = dateColumn + columnOffset
    If dateColumn < 1 Or dateColumn > targetSheet.Columns.Count Then
        GetWorkloadByDate = CVErr(xlErrRef) 'Offset out of bounds
        Exit Function
    End If

    '--- Return result from last data row + 1
    GetWorkloadByDate = targetSheet.Cells(lastDataRow + 1, dateColumn).value
    Exit Function

ErrorHandler:
    GetWorkloadByDate = CVErr(xlErrValue)
End Function

'@Description("Counts absent employees on a given date (based on absence codes)")
'@Param targetDate Date to check
'@Returns Count of absent employees
Public Function CountAbsentEmployees(ByVal targetDate As Date) As Long
    Dim dateColumn As Long
    dateColumn = DateHelpers.FindDateColumn(Tabelle3, 10, targetDate, 15)

    If dateColumn = 0 Then
        CountAbsentEmployees = 0
        Exit Function
    End If

    Dim targetRange As Range
    Set targetRange = Intersect(Tabelle3.Columns(dateColumn), Tabelle3.UsedRange)

    CountAbsentEmployees = CountAbsenceCodes(targetRange)
End Function

'@Description("Counts cells containing specific absence codes (F, U, K, WK, S, ÜK, T)")
'@Param targetRange Range to scan
'@Returns Count of cells with absence codes
Public Function CountAbsenceCodes(ByVal targetRange As Range) As Long
    '@Ignore VariableNotUsed
    Dim currentCell As Range
    Dim matchCount As Long
    Dim absenceCodesArray As Variant
    Dim codeIndex As Long

    '--- Define codes to count
    absenceCodesArray = Array("F", "U", "K", "WK", "S", "ÜK", "T")

    matchCount = 0

    '--- PERFORMANCE: Direct cell iteration (fast for columnar data)
    For Each currentCell In targetRange.Cells
        If Not IsEmpty(currentCell.value) Then
            For codeIndex = LBound(absenceCodesArray) To UBound(absenceCodesArray)
                If StrComp(Trim$(CStr(currentCell.value)), absenceCodesArray(codeIndex), vbTextCompare) = 0 Then
                    matchCount = matchCount + 1
                    Exit For 'Don't count twice
                End If
            Next codeIndex
        End If
    Next currentCell

    CountAbsenceCodes = matchCount
End Function

'@Description("Calculates workload: percentage of available employees (excluding specified absence codes)")
'@Param exclusionRange Range containing absence codes to exclude (e.g., F, U, K)
'@Param includeDepartment Whether to include all employees or only visible rows
'@Returns Workload ratio (0.0 to 1.0)
Public Function CalculateWorkload( _
        ByVal exclusionRange As Range, _
        Optional ByVal includeDepartment As Boolean = False) As Double

    Application.Volatile True

    On Error GoTo ErrorHandler

    Dim callerSheet As Worksheet
    Set callerSheet = Application.Caller.Worksheet

    Dim dataTable As ListObject
    Set dataTable = callerSheet.ListObjects(1)

    '--- Find column based on formula position
    Dim columnIndex As Long
    columnIndex = Application.Caller.Column - dataTable.Range.Columns(1).Column + 1

    If columnIndex < 1 Or columnIndex > dataTable.ListColumns.Count Then GoTo SafeExit

    Dim columnName As String
    columnName = dataTable.HeaderRowRange.Cells(1, columnIndex).value

    Dim dayRange As Range
    Set dayRange = dataTable.ListColumns(columnName).DataBodyRange

    Dim employeeRange As Range
    Set employeeRange = dataTable.ListColumns("Mitarbeiter").DataBodyRange

    '--- Calculate based on mode
    If includeDepartment Then
        CalculateWorkload = CalculateWorkloadAllRows(dayRange, employeeRange, exclusionRange)
    Else
        CalculateWorkload = CalculateWorkloadVisibleRows(dayRange, employeeRange, exclusionRange)
    End If

    Exit Function

SafeExit:
    CalculateWorkload = 0#
    Exit Function

ErrorHandler:
    Resume SafeExit
End Function

'@Description("Counts available employees (excluding specified absence codes)")
'@Param exclusionRange Range containing absence codes to exclude
'@Param includeDepartment Whether to include all employees or only visible rows
'@Returns Count of available employees
Public Function CountAvailableEmployees( _
        ByVal exclusionRange As Range, _
        Optional ByVal includeDepartment As Boolean = False) As Long

    Application.Volatile True

    On Error GoTo ErrorHandler

    Dim callerSheet As Worksheet
    Set callerSheet = Application.Caller.Worksheet

    Dim dataTable As ListObject
    Set dataTable = callerSheet.ListObjects(1)

    Dim columnIndex As Long
    columnIndex = Application.Caller.Column - dataTable.Range.Columns(1).Column + 1

    If columnIndex < 1 Or columnIndex > dataTable.ListColumns.Count Then GoTo SafeExit

    Dim columnName As String
    columnName = dataTable.HeaderRowRange.Cells(1, columnIndex).value

    Dim dayRange As Range
    Set dayRange = dataTable.ListColumns(columnName).DataBodyRange

    Dim employeeRange As Range
    Set employeeRange = dataTable.ListColumns("Mitarbeiter").DataBodyRange

    If includeDepartment Then
        CountAvailableEmployees = CountAvailableAllRows(dayRange, employeeRange, exclusionRange)
    Else
        CountAvailableEmployees = CountAvailableVisibleRows(dayRange, employeeRange, exclusionRange)
    End If

    Exit Function

SafeExit:
    CountAvailableEmployees = 0
    Exit Function

ErrorHandler:
    Resume SafeExit
End Function

'--- PRIVATE HELPER FUNCTIONS ---

'@Description("Calculates workload for visible rows only")
Private Function CalculateWorkloadVisibleRows( _
        ByVal dayRange As Range, _
        ByVal employeeRange As Range, _
        ByVal exclusionRange As Range) As Double

    Dim exclusionDict As Dictionary
    Set exclusionDict = BuildExclusionDictionary(exclusionRange)

    Dim rowIndex As Long
    Dim dayCell As Range
    Dim employeeCell As Range
    Dim cellValue As String
    Dim availableCount As Long
    Dim totalCount As Long

    '--- PERFORMANCE: Iterate visible rows only
    For rowIndex = 1 To dayRange.Rows.Count
        Set dayCell = dayRange.Cells(rowIndex, 1)
        Set employeeCell = employeeRange.Cells(rowIndex, 1)

        If Not dayCell.EntireRow.Hidden Then
            If EmployeeService.IsCellNotEmpty(employeeCell) Then
                cellValue = Trim$(SafeCellString(dayCell.Value2))

                If Len(cellValue) > 0 And Not exclusionDict.Exists(cellValue) Then
                    availableCount = availableCount + 1
                    totalCount = totalCount + 1
                ElseIf Len(cellValue) = 0 Then
                    totalCount = totalCount + 1
                End If
            End If
        End If
    Next rowIndex

    If totalCount = 0 Then
        CalculateWorkloadVisibleRows = 0#
    Else
        CalculateWorkloadVisibleRows = availableCount / totalCount
    End If
End Function

'@Description("Calculates workload for all rows (including hidden)")
Private Function CalculateWorkloadAllRows( _
        ByVal dayRange As Range, _
        ByVal employeeRange As Range, _
        ByVal exclusionRange As Range) As Double

    Dim exclusionDict As Dictionary
    Set exclusionDict = BuildExclusionDictionary(exclusionRange)

    Dim rowIndex As Long
    Dim dayCell As Range
    Dim employeeCell As Range
    Dim cellValue As String
    Dim availableCount As Long
    Dim totalCount As Long

    For rowIndex = 1 To dayRange.Rows.Count
        Set dayCell = dayRange.Cells(rowIndex, 1)
        Set employeeCell = employeeRange.Cells(rowIndex, 1)

        If EmployeeService.IsCellNotEmpty(employeeCell) Then
            cellValue = Trim$(SafeCellString(dayCell.Value2))

            If Len(cellValue) > 0 And Not exclusionDict.Exists(cellValue) Then
                availableCount = availableCount + 1
                totalCount = totalCount + 1
            ElseIf Len(cellValue) = 0 Then
                totalCount = totalCount + 1
            End If
        End If
    Next rowIndex

    If totalCount = 0 Then
        CalculateWorkloadAllRows = 0#
    Else
        CalculateWorkloadAllRows = availableCount / totalCount
    End If
End Function

'@Description("Counts available employees in visible rows")
Private Function CountAvailableVisibleRows( _
        ByVal dayRange As Range, _
        ByVal employeeRange As Range, _
        ByVal exclusionRange As Range) As Long

    Dim exclusionDict As Dictionary
    Set exclusionDict = BuildExclusionDictionary(exclusionRange)

    Dim rowIndex As Long
    Dim dayCell As Range
    Dim employeeCell As Range
    Dim cellValue As String
    Dim availableCount As Long

    For rowIndex = 1 To dayRange.Rows.Count
        Set dayCell = dayRange.Cells(rowIndex, 1)
        Set employeeCell = employeeRange.Cells(rowIndex, 1)

        If Not dayCell.EntireRow.Hidden Then
            If EmployeeService.IsCellNotEmpty(employeeCell) Then
                cellValue = Trim$(SafeCellString(dayCell.Value2))
                If Len(cellValue) = 0 Then
                    availableCount = availableCount + 1
                End If
            End If
        End If
    Next rowIndex

    CountAvailableVisibleRows = availableCount
End Function

'@Description("Counts available employees in all rows")
Private Function CountAvailableAllRows( _
        ByVal dayRange As Range, _
        ByVal employeeRange As Range, _
        ByVal exclusionRange As Range) As Long

    Dim exclusionDict As Dictionary
    Set exclusionDict = BuildExclusionDictionary(exclusionRange)

    Dim rowIndex As Long
    Dim dayCell As Range
    Dim employeeCell As Range
    Dim cellValue As String
    Dim availableCount As Long

    For rowIndex = 1 To dayRange.Rows.Count
        Set dayCell = dayRange.Cells(rowIndex, 1)
        Set employeeCell = employeeRange.Cells(rowIndex, 1)

        If EmployeeService.IsCellNotEmpty(employeeCell) Then
            cellValue = Trim$(SafeCellString(dayCell.Value2))
            If Len(cellValue) = 0 Or Not exclusionDict.Exists(cellValue) Then
                availableCount = availableCount + 1
            End If
        End If
    Next rowIndex

    CountAvailableAllRows = availableCount
End Function

'@Description("Builds exclusion dictionary from range (case-insensitive)")
Private Function BuildExclusionDictionary(ByVal exclusionRange As Range) As Dictionary
    Dim exclusionDict As Dictionary
    Set exclusionDict = New Dictionary
    exclusionDict.CompareMode = vbTextCompare

    Dim currentCell As Range
    Dim keyValue As String

    For Each currentCell In exclusionRange.Cells
        If Not IsError(currentCell.Value2) Then
            keyValue = Trim$(CStr(currentCell.Value2))
            If Len(keyValue) > 0 Then
                If Not exclusionDict.Exists(keyValue) Then
                    exclusionDict.Add keyValue, True
                End If
            End If
        End If
    Next currentCell

    Set BuildExclusionDictionary = exclusionDict
End Function

'@Description("Safely converts cell value to string")
Private Function SafeCellString(ByVal cellValue As Variant) As String
    On Error Resume Next
    If IsError(cellValue) Or IsNull(cellValue) Or IsEmpty(cellValue) Then
        SafeCellString = vbNullString
    Else
        SafeCellString = CStr(cellValue)
    End If
End Function

'@Description("Gets last used column in sheet")
Private Function GetSheetLastUsedColumn(ByVal targetSheet As Worksheet) As Long
    On Error GoTo Fail
    GetSheetLastUsedColumn = targetSheet.Columns.Count
    Exit Function
Fail:
    GetSheetLastUsedColumn = 0
End Function

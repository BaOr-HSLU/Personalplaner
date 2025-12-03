Attribute VB_Name = "LinkingService"
'@Folder("Services.Linking")
'@ModuleDescription("Bidirectional synchronization between Personalplaner (overview) and KW sheets (weekly view)")
Option Explicit

'--- Prevent infinite loops during sync
Private syncInProgress As Boolean

'@Description("Maps absence code long form to short form (e.g., 'Ferien' -> 'F')")
'@Param longForm The long form text (e.g., 'Ferien', 'Krank')
'@Returns Short form code (e.g., 'F', 'K') or original value if not an absence code
Private Function MapAbsenceCodeToShortForm(ByVal longForm As String) As String
    If Len(Trim$(longForm)) = 0 Then
        MapAbsenceCodeToShortForm = vbNullString
        Exit Function
    End If

    Dim absenceCodes As Dictionary
    Set absenceCodes = AbsenceCode.GetAllCodes

    Dim codeKey As Variant
    Dim currentCode As AbsenceCode

    For Each codeKey In absenceCodes.Keys
        Set currentCode = absenceCodes(codeKey)

        If UCase(Trim$(longForm)) = UCase(currentCode.LongForm) Then
            MapAbsenceCodeToShortForm = currentCode.ShortForm
            Exit Function
        End If
    Next codeKey

    '--- Not an absence code, return original value (e.g., project name)
    MapAbsenceCodeToShortForm = longForm
End Function

'@Description("Maps absence code short form to long form (e.g., 'F' -> 'Ferien')")
'@Param shortForm The short form code (e.g., 'F', 'K')
'@Returns Long form text (e.g., 'Ferien', 'Krank') or original value if not an absence code
Private Function MapAbsenceCodeToLongForm(ByVal shortForm As String) As String
    If Len(Trim$(shortForm)) = 0 Then
        MapAbsenceCodeToLongForm = vbNullString
        Exit Function
    End If

    Dim absenceCodes As Dictionary
    Set absenceCodes = AbsenceCode.GetAllCodes

    Dim codeKey As Variant
    Dim currentCode As AbsenceCode

    For Each codeKey In absenceCodes.Keys
        Set currentCode = absenceCodes(codeKey)

        If UCase(Trim$(shortForm)) = UCase(currentCode.ShortForm) Then
            MapAbsenceCodeToLongForm = currentCode.LongForm
            Exit Function
        End If
    Next codeKey

    '--- Not an absence code, return original value (e.g., project name)
    MapAbsenceCodeToLongForm = shortForm
End Function

'@Description("Synchronizes a cell change from KW sheet to Personalplaner overview")
'@Param kwSheet The weekly (KW) sheet where the change occurred
'@Param changedCell The cell that was changed
Public Sub SyncCellToOverview(ByVal kwSheet As Worksheet, ByVal changedCell As Range)
    On Error GoTo ErrorHandler

    '--- Prevent infinite loop
    If syncInProgress Then Exit Sub
    syncInProgress = True

    '--- Only sync data in ListObject (not headers or other areas)
    Dim parentTable As ListObject
    On Error Resume Next
    Set parentTable = changedCell.ListObject
    On Error GoTo ErrorHandler

    If parentTable Is Nothing Then GoTo CleanupAndExit

    '--- Only sync columns E to I (5 weekdays: Monday-Friday)
    Dim changedColumnIndex As Long
    changedColumnIndex = changedCell.Column - parentTable.Range.Column + 1

    If changedColumnIndex < 5 Or changedColumnIndex > 9 Then GoTo CleanupAndExit

    '--- Get employee number from column 1 (key for matching)
    Dim employeeNumber As String
    Dim changedRow As ListRow
    Set changedRow = changedCell.ListObject.ListRows(changedCell.Row - changedCell.ListObject.Range.Row)
    employeeNumber = CStr(changedRow.Range.Cells(1, 1).value)

    If Len(Trim$(employeeNumber)) = 0 Then GoTo CleanupAndExit

    '--- Calculate which weekday (0=Monday, 4=Friday)
    Dim weekdayOffset As Long
    weekdayOffset = changedColumnIndex - 5  '--- Column E=0, F=1, G=2, H=3, I=4

    '--- Parse KW number from sheet name (e.g., "KW49 2025")
    Dim kwNumber As Long
    Dim yearValue As Long
    Call ParseKWSheetName(kwSheet.Name, kwNumber, yearValue)

    If kwNumber = 0 Then GoTo CleanupAndExit

    '--- Calculate the actual date for this weekday in this KW
    Dim targetDate As Date
    targetDate = GetDateFromKWAndWeekday(yearValue, kwNumber, weekdayOffset)

    '--- Find the corresponding cell in Personalplaner
    Dim overviewSheet As Worksheet
    Set overviewSheet = Tabelle3

    '--- Find employee row in overview by employee number
    Dim overviewRow As ListRow
    Set overviewRow = FindEmployeeRowByNumber(overviewSheet, employeeNumber)

    If overviewRow Is Nothing Then GoTo CleanupAndExit

    '--- Find date column in overview
    Dim dateColumn As Long
    dateColumn = DateHelpers.FindDateColumn(overviewSheet, 10, targetDate, 15)

    If dateColumn = 0 Then GoTo CleanupAndExit

    '--- Transform value: LongForm -> ShortForm (e.g., "Ferien" -> "F")
    Dim originalValue As String
    Dim transformedValue As String
    originalValue = CStr(changedCell.value)
    transformedValue = MapAbsenceCodeToShortForm(originalValue)

    '--- Write to overview (disable events to prevent loop)
    Application.EnableEvents = False

    Dim targetCell As Range
    Set targetCell = overviewRow.Range.Cells(1, dateColumn - overviewRow.Range.Parent.Cells(overviewRow.Range.Row, 1).Column + 1)
    targetCell.value = transformedValue

    Application.EnableEvents = True

CleanupAndExit:
    syncInProgress = False
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    syncInProgress = False
    Debug.Print "LinkingService.SyncCellToOverview Error: " & Err.Description
End Sub

'@Description("Synchronizes a cell change from Personalplaner overview to KW sheet")
'@Param overviewSheet The Personalplaner sheet where the change occurred
'@Param changedCell The cell that was changed
Public Sub SyncCellToWeeklySheet(ByVal overviewSheet As Worksheet, ByVal changedCell As Range)
    On Error GoTo ErrorHandler

    '--- Prevent infinite loop
    If syncInProgress Then Exit Sub
    syncInProgress = True

    '--- Only sync data in ListObject (not headers or other areas)
    Dim parentTable As ListObject
    On Error Resume Next
    Set parentTable = changedCell.ListObject
    On Error GoTo ErrorHandler

    If parentTable Is Nothing Then GoTo CleanupAndExit

    '--- Only sync columns >= 15 (O+: calendar columns)
    Dim changedColumnIndex As Long
    changedColumnIndex = changedCell.Column

    If changedColumnIndex < 15 Then GoTo CleanupAndExit

    '--- Get employee number from column 6 (Number column in table, absolute column F)
    Dim employeeNumber As String
    Dim changedRow As ListRow
    Set changedRow = changedCell.ListObject.ListRows(changedCell.Row - changedCell.ListObject.Range.Row)
    employeeNumber = CStr(changedRow.Range.Cells(1, 6).value)  '--- Column 6 = "Number"

    If Len(Trim$(employeeNumber)) = 0 Then GoTo CleanupAndExit

    '--- Get date from column header (row 10)
    Dim targetDate As Date
    On Error Resume Next
    targetDate = overviewSheet.Cells(10, changedColumnIndex).value
    On Error GoTo ErrorHandler

    If targetDate = 0 Then GoTo CleanupAndExit

    '--- Calculate KW and weekday from date
    Dim kwNumber As Long
    Dim yearValue As Long
    Dim weekdayOffset As Long

    kwNumber = WorksheetFunction.WeekNum(targetDate, 2)  '--- ISO week number
    yearValue = Year(targetDate)
    weekdayOffset = Weekday(targetDate, vbMonday) - 1  '--- 0=Monday, 4=Friday

    '--- Find corresponding KW sheet (e.g., "KW49 2025")
    Dim kwSheetName As String
    kwSheetName = "KW" & kwNumber & " " & yearValue

    Dim kwSheet As Worksheet
    On Error Resume Next
    Set kwSheet = ThisWorkbook.Worksheets(kwSheetName)
    On Error GoTo ErrorHandler

    If kwSheet Is Nothing Then GoTo CleanupAndExit  '--- KW sheet doesn't exist yet

    '--- Find employee row in KW sheet by employee number
    Dim kwRow As ListRow
    Set kwRow = FindEmployeeRowByNumber(kwSheet, employeeNumber)

    If kwRow Is Nothing Then GoTo CleanupAndExit

    '--- Calculate target column in KW sheet (E=Monday, F=Tuesday, ..., I=Friday)
    Dim targetColumnIndex As Long
    targetColumnIndex = 5 + weekdayOffset  '--- E=5, F=6, G=7, H=8, I=9

    '--- Transform value: ShortForm -> LongForm (e.g., "F" -> "Ferien")
    Dim originalValue As String
    Dim transformedValue As String
    originalValue = CStr(changedCell.value)
    transformedValue = MapAbsenceCodeToLongForm(originalValue)

    '--- Write to KW sheet (disable events to prevent loop)
    Application.EnableEvents = False

    Dim targetCell As Range
    Set targetCell = kwRow.Range.Cells(1, targetColumnIndex)
    targetCell.value = transformedValue

    Application.EnableEvents = True

CleanupAndExit:
    syncInProgress = False
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    syncInProgress = False
    Debug.Print "LinkingService.SyncCellToWeeklySheet Error: " & Err.Description
End Sub

'@Description("Finds an employee row in a ListObject by employee number")
'@Param targetSheet The worksheet to search
'@Param employeeNumber The employee number to find
'@Returns The ListRow object or Nothing if not found
Private Function FindEmployeeRowByNumber(ByVal targetSheet As Worksheet, ByVal employeeNumber As String) As ListRow
    On Error Resume Next

    Dim searchTable As ListObject
    Set searchTable = targetSheet.ListObjects(1)  '--- First table on sheet

    If searchTable Is Nothing Then
        Set FindEmployeeRowByNumber = Nothing
        Exit Function
    End If

    Dim currentRow As ListRow
    Dim currentNumber As String

    For Each currentRow In searchTable.ListRows
        currentNumber = CStr(currentRow.Range.Cells(1, 1).value)

        If currentNumber = employeeNumber Then
            Set FindEmployeeRowByNumber = currentRow
            Exit Function
        End If
    Next currentRow

    Set FindEmployeeRowByNumber = Nothing
End Function

'@Description("Parses KW sheet name to extract week number and year")
'@Param sheetName The sheet name (e.g., 'KW49 2025')
'@Param kwNumber Output: The KW number
'@Param yearValue Output: The year
Private Sub ParseKWSheetName(ByVal sheetName As String, ByRef kwNumber As Long, ByRef yearValue As Long)
    On Error Resume Next

    '--- Example: "KW49 2025"
    Dim parts() As String
    parts = Split(sheetName, " ")

    If UBound(parts) >= 1 Then
        kwNumber = CLng(Mid$(parts(0), 3))  '--- Remove "KW" prefix
        yearValue = CLng(parts(1))
    Else
        kwNumber = 0
        yearValue = 0
    End If
End Sub

'@Description("Calculates the date for a specific weekday in a given ISO week")
'@Param yearValue The year
'@Param kwNumber The ISO week number
'@Param weekdayOffset The weekday offset (0=Monday, 4=Friday)
'@Returns The calculated date
Private Function GetDateFromKWAndWeekday(ByVal yearValue As Long, ByVal kwNumber As Long, ByVal weekdayOffset As Long) As Date
    '--- Calculate Monday of the ISO week
    Dim jan4 As Date
    Dim mondayOfWeek1 As Date
    Dim weekday_jan4 As Long

    jan4 = DateSerial(yearValue, 1, 4)
    weekday_jan4 = Weekday(jan4, vbMonday)
    mondayOfWeek1 = jan4 - (weekday_jan4 - 1)

    '--- Calculate Monday of requested week
    Dim mondayOfTargetWeek As Date
    mondayOfTargetWeek = mondayOfWeek1 + (kwNumber - 1) * 7

    '--- Add weekday offset
    GetDateFromKWAndWeekday = mondayOfTargetWeek + weekdayOffset
End Function

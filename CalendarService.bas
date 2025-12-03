Attribute VB_Name = "CalendarService"
'@Folder("Services.Calendar")
'@ModuleDescription("Creates and manages calendar sheets with work days, holidays, and formatting")
Option Explicit

Private Const EMPLOYEE_ROWS_COUNT As Long = 50
Private Const DATE_ROW_OFFSET As Long = -1
Private Const CALENDAR_WEEK_ROW_OFFSET As Long = -2
Private Const MONTH_ROW_OFFSET As Long = -3
Private Const HOLIDAYS_ROW_OFFSET As Long = -5
Private Const VACATIONS_ROW_OFFSET As Long = -4

'--- Color constant for holidays and vacations (light blue)
Private Const HOLIDAY_VACATION_COLOR As Long = 15849925  ' RGB(165, 205, 241)

'@Description("Creates a calendar with work days only (Monday-Friday)")
'@Param startCell The cell where the calendar should start
Public Sub CreateWorkDayCalendar(ByVal startCell As Range)
    If startCell Is Nothing Then
        MsgBox "Keine Startzelle ausgewaehlt.", vbExclamation
        Exit Sub
    End If

    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet

    '--- Show UserForm to get options
    Dim optionsForm As UF_CalendarOptions
    Set optionsForm = New UF_CalendarOptions

    optionsForm.Show

    '--- Check if user cancelled
    If optionsForm.IsCancelled Then
        Unload optionsForm
        Exit Sub
    End If

    '--- Get values from form
    Dim startDate As Date
    Dim endDate As Date
    Dim IncludeHolidaysVacations As Boolean
    Dim ApplyConditionalFormatting As Boolean
    Dim CopyFormattingFromA2 As Boolean

    startDate = optionsForm.startDate
    endDate = optionsForm.endDate
    IncludeHolidaysVacations = optionsForm.IncludeHolidaysVacations
    ApplyConditionalFormatting = optionsForm.ApplyConditionalFormatting
    CopyFormattingFromA2 = optionsForm.CopyFormattingFromA2

    Unload optionsForm

    '--- FIX #5: Clear existing calendar elements before creating new one
    Call ClearExistingCalendar(targetSheet, startCell)

    '--- Performance optimization
    Dim originalScreenUpdating As Boolean
    Dim originalCursor As XlMousePointer
    Dim originalCalculation As XlCalculation

    originalScreenUpdating = Application.ScreenUpdating
    originalCursor = Application.Cursor
    originalCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Application.Calculation = xlCalculationManual

    Dim currentColumn As Long
    Dim currentRow As Long

    currentColumn = startCell.Column
    currentRow = startCell.Row

    Dim currentDate As Date
    Dim currentCalendarWeek As Long
    Dim currentMonth As String
    Dim currentYear As String

    currentDate = startDate
    currentCalendarWeek = WorksheetFunction.WeekNum(currentDate, 2)
    currentMonth = Format(currentDate, "MMMM")
    currentYear = Format(currentDate, "YYYY")

    Dim weekStartColumn As Long
    Dim monthStartColumn As Long
    Dim yearStartColumn As Long

    weekStartColumn = currentColumn
    monthStartColumn = currentColumn
    yearStartColumn = currentColumn

    '--- Draw left border for first column
    With targetSheet.Range(targetSheet.Cells(currentRow + DATE_ROW_OFFSET, currentColumn), _
                           targetSheet.Cells(currentRow + EMPLOYEE_ROWS_COUNT, currentColumn)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With

    Dim firstCalendarColumn As Long
    firstCalendarColumn = currentColumn

    '--- Main loop: iterate through all dates
    Do While currentDate <= endDate
        '--- Only process weekdays (Monday-Friday)
        If Weekday(currentDate, vbMonday) <= 5 Then
            Application.StatusBar = currentYear & " / " & currentMonth & " / " & currentCalendarWeek & " / " & currentDate

            '--- Store actual date as value with weekday format (Mo, Di, Mi, Do, Fr)
            targetSheet.Cells(currentRow, currentColumn).value = currentDate
            targetSheet.Cells(currentRow, currentColumn).NumberFormat = "ddd"  '--- German short weekday format
            targetSheet.Cells(currentRow, currentColumn).HorizontalAlignment = xlCenter
            targetSheet.Cells(currentRow, currentColumn).Font.Bold = True
            targetSheet.Cells(currentRow, currentColumn).Font.Size = 8

            '--- FIX #8: Increase column width to 2.0
            targetSheet.Columns(currentColumn).ColumnWidth = 2#

            '--- Format holiday/vacation cell (merged vertical)
            With targetSheet.Range(targetSheet.Cells(currentRow - 5, currentColumn), _
                                   targetSheet.Cells(currentRow - 8, currentColumn))
                .Merge
                .Font.Size = 6
                .Orientation = 90
                .VerticalAlignment = xlBottom
                .HorizontalAlignment = xlCenter
            End With

            '--- FIX #6: Add dotted border between individual days
            If currentColumn > firstCalendarColumn Then
                With targetSheet.Range(targetSheet.Cells(currentRow, currentColumn), _
                                       targetSheet.Cells(currentRow + EMPLOYEE_ROWS_COUNT, currentColumn)).Borders(xlEdgeLeft)
                    .LineStyle = xlDot
                    .Weight = xlThin
                    .Color = RGB(192, 192, 192)
                End With
            End If

            '--- Check for calendar week change
            If WorksheetFunction.WeekNum(currentDate, 2) <> currentCalendarWeek Then
                Call FinalizeCalendarWeek(targetSheet, currentRow, weekStartColumn, currentColumn - 1, currentCalendarWeek)

                '--- Draw solid border for new week (overrides dotted border)
                With targetSheet.Range(targetSheet.Cells(currentRow, currentColumn), _
                                       targetSheet.Cells(currentRow + EMPLOYEE_ROWS_COUNT, currentColumn)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(0, 0, 0)
                End With

                currentCalendarWeek = WorksheetFunction.WeekNum(currentDate, 2)
                weekStartColumn = currentColumn
            End If

            '--- Check for month change
            If Format(currentDate, "MMMM") <> currentMonth Then
                Call FinalizeMonth(targetSheet, currentRow, monthStartColumn, currentColumn - 1, currentMonth, currentYear)
                currentMonth = Format(currentDate, "MMMM")
                monthStartColumn = currentColumn
            End If

            '--- Check for year change
            If Format(currentDate, "YYYY") <> currentYear Then
                currentYear = Format(currentDate, "YYYY")
            End If

            currentColumn = currentColumn + 1
        End If

        currentDate = currentDate + 1
    Loop

    '--- Finalize last week
    Call FinalizeCalendarWeek(targetSheet, currentRow, weekStartColumn, currentColumn - 1, currentCalendarWeek)

    '--- Finalize last month
    Call FinalizeMonth(targetSheet, currentRow, monthStartColumn, currentColumn - 1, currentMonth, currentYear)

    '--- Create named range for dates
    On Error Resume Next
    ThisWorkbook.Names("TAGE").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:="TAGE", _
        RefersTo:=targetSheet.Range(targetSheet.Cells(startCell.Row, startCell.Column), _
                                     targetSheet.Cells(startCell.Row, currentColumn - 1))

    '--- FIX #4: Extend ListObjects to include new calendar columns
    Call ExtendListObjectsToCalendar(targetSheet, startCell.Row, firstCalendarColumn, currentColumn - 1)

    Application.ScreenUpdating = originalScreenUpdating
    Application.Cursor = originalCursor
    Application.Calculation = originalCalculation
    Application.StatusBar = False

    'MsgBox "Kalender mit Arbeitstagen erfolgreich erstellt!", vbInformation

    '--- Apply options from UserForm
    If IncludeHolidaysVacations Then
        Call AddHolidaysAndVacations
    End If

    If CopyFormattingFromA2 Then
        Call CopyConditionalFormatting(targetSheet, startCell)
    End If

    If ApplyConditionalFormatting Then
        '--- FIX #1 & #2: Apply conditional formatting AND data validation dropdowns
        Call ApplyConditionalFormattingToTables
        Call ApplyDataValidationToTables
    End If

    '--- FIX: Stay on Personalplaner instead of switching to Tabelle1
    Tabelle3.Activate
End Sub

'@Description("Clears existing calendar elements to avoid conflicts")
'@Param targetSheet The worksheet containing the calendar
'@Param startCell The cell where the calendar starts
Private Sub ClearExistingCalendar(ByVal targetSheet As Worksheet, ByVal startCell As Range)
    On Error Resume Next

    '--- Delete named range
    ThisWorkbook.Names("TAGE").Delete

    '--- Clear calendar area (from start cell to reasonable extent)
    '--- Clear header rows (5 rows above start cell)
    '--- Clear data rows (50 rows below start cell)
    Dim clearRange As Range
    Set clearRange = targetSheet.Range( _
        targetSheet.Cells(startCell.Row - 8, startCell.Column), _
        targetSheet.Cells(startCell.Row + EMPLOYEE_ROWS_COUNT, 300))

    '--- Clear contents, formats, and validation
    clearRange.ClearContents
    clearRange.ClearFormats
    clearRange.Validation.Delete

    On Error GoTo 0
End Sub

'@Description("Extends all ListObjects on the sheet to include calendar columns")
'@Param targetSheet The worksheet containing the tables
'@Param dataRow The row where employee data starts (not used, kept for compatibility)
'@Param firstColumn The first calendar column
'@Param lastColumn The last calendar column
Private Sub ExtendListObjectsToCalendar(ByVal targetSheet As Worksheet, _
                                         ByVal dataRow As Long, _
                                         ByVal firstColumn As Long, _
                                         ByVal lastColumn As Long)
    On Error Resume Next

    Dim listObj As ListObject
    For Each listObj In targetSheet.ListObjects
        '--- FIX: Extend ALL tables on the sheet to the last calendar column
        '--- This ensures that all employee tables include the full calendar
        If Not listObj.DataBodyRange Is Nothing Then
            '--- Calculate new range: from table start to last calendar column
            Dim newRange As Range
            Dim lastRow As Long
            lastRow = listObj.Range.Row + listObj.Range.Rows.Count - 1

            Set newRange = targetSheet.Range( _
                listObj.Range.Cells(1, 1), _
                targetSheet.Cells(lastRow, lastColumn))

            '--- Only resize if the new range is actually larger
            If newRange.Columns.Count > listObj.Range.Columns.Count Then
                listObj.Resize newRange
            End If
        End If
    Next listObj

    On Error GoTo 0
End Sub

'@Description("Finalizes a calendar week by merging cells and adding borders")
Private Sub FinalizeCalendarWeek(ByVal targetSheet As Worksheet, _
                                  ByVal dataRow As Long, _
                                  ByVal startColumn As Long, _
                                  ByVal endColumn As Long, _
                                  ByVal weekNumber As Long)

    '--- Calendar week number
    With targetSheet.Range(targetSheet.Cells(dataRow + CALENDAR_WEEK_ROW_OFFSET, startColumn), _
                           targetSheet.Cells(dataRow + CALENDAR_WEEK_ROW_OFFSET, endColumn))
        .Merge
        .value = CStr(weekNumber)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 10
        '--- Only set outer borders, not inner vertical lines
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
    End With

    '--- Date range (e.g., "01-05")
    Dim firstDayDate As Date
    Dim lastDayDate As Date
    Dim currentYear As Long
    
    '--- Determine the year (use current year or year from context)
    currentYear = Year(Date)
    
    '--- Calculate the Monday of the given ISO week number
    firstDayDate = GetMondayOfWeek(currentYear, weekNumber)
    lastDayDate = firstDayDate + 4  ' Monday to Friday

    With targetSheet.Range(targetSheet.Cells(dataRow + DATE_ROW_OFFSET, startColumn), _
                           targetSheet.Cells(dataRow + DATE_ROW_OFFSET, endColumn))
        .Merge
        .NumberFormat = "@"
        .value = Format(firstDayDate, "dd") & "-" & Format(lastDayDate, "dd")
        .HorizontalAlignment = xlCenter
        .Font.Bold = False
        .Font.Size = 8
        '--- Only set outer borders, not inner vertical lines
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
    End With
End Sub

'--- Helper function to get the Monday of a specific ISO week
Private Function GetMondayOfWeek(ByVal yearNumber As Long, ByVal weekNumber As Long) As Date
    Dim jan4 As Date
    Dim mondayOfWeek1 As Date
    Dim weekday_jan4 As Long
    
    '--- ISO 8601: Week 1 is the week containing January 4th
    jan4 = DateSerial(yearNumber, 1, 4)
    
    '--- Find the Monday of week 1
    weekday_jan4 = Weekday(jan4, vbMonday)  ' 1=Monday, 7=Sunday
    mondayOfWeek1 = jan4 - (weekday_jan4 - 1)
    
    '--- Calculate the Monday of the requested week
    GetMondayOfWeek = mondayOfWeek1 + (weekNumber - 1) * 7
End Function

'@Description("Copies conditional formatting from cell A2 to all calendar week cells")
'@Param targetSheet The worksheet containing the calendar
'@Param startCell The cell where the calendar starts
Private Sub CopyConditionalFormatting(ByVal targetSheet As Worksheet, ByVal startCell As Range)
    On Error Resume Next
    Dim sourceCell As Range
    Set sourceCell = targetSheet.Range("A2")
    
    '--- Check if source cell has conditional formatting
    If sourceCell.FormatConditions.Count = 0 Then
        MsgBox "Zelle A2 hat keine bedingte Formatierung zum Kopieren.", vbInformation
        Exit Sub
    End If
    
    '--- Find all week cells in the calendar
    Dim datesRange As Range
    Set datesRange = targetSheet.Range("TAGE")
    
    If datesRange Is Nothing Then
        MsgBox "Kalender-Bereich 'TAGE' wurde nicht gefunden.", vbExclamation
        Exit Sub
    End If
    
    Dim weekRow As Long
    weekRow = datesRange.Row + CALENDAR_WEEK_ROW_OFFSET
    
    Dim firstColumn As Long
    Dim lastColumn As Long
    firstColumn = datesRange.Column
    lastColumn = datesRange.Column + datesRange.Columns.Count - 1
    
    Application.ScreenUpdating = False
    
    '--- Collect all week cell ranges using Union
    Dim allWeekRanges As Range
    Set allWeekRanges = Nothing
    
    Dim currentColumn As Long
    currentColumn = firstColumn
    
    Do While currentColumn <= lastColumn
        Dim weekCell As Range
        Set weekCell = targetSheet.Cells(weekRow, currentColumn)
        
        '--- If it's a merged cell, add the merge area
        If weekCell.MergeCells Then
            Dim weekCellArea As Range
            Set weekCellArea = weekCell.MergeArea
            
            '--- Add to union
            If allWeekRanges Is Nothing Then
                Set allWeekRanges = weekCellArea
            Else
                Set allWeekRanges = Application.Union(allWeekRanges, weekCellArea)
            End If
            
            '--- Move to next week (skip merged cells)
            currentColumn = weekCellArea.Column + weekCellArea.Columns.Count
        Else
            '--- Add single cell to union
            If allWeekRanges Is Nothing Then
                Set allWeekRanges = weekCell
            Else
                Set allWeekRanges = Application.Union(allWeekRanges, weekCell)
            End If
            currentColumn = currentColumn + 1
        End If
    Loop
    
    '--- For each conditional format in A2, extend its AppliesTo range
    Dim fc As FormatCondition
    Dim i As Long
    Dim combinedRange As Range
    
    For i = 1 To sourceCell.FormatConditions.Count
        Set fc = sourceCell.FormatConditions(i)
        
        '--- Combine original range with all week ranges using Union
        On Error Resume Next
        Set combinedRange = Application.Union(fc.AppliesTo, allWeekRanges)
        On Error GoTo 0
        
        If Not combinedRange Is Nothing Then
            Debug.Print "--- Format Condition", i
            Debug.Print "Original:", fc.AppliesTo.Address
            Debug.Print "Combined:", combinedRange.Address
            
            '--- Extend the AppliesTo property
            On Error Resume Next
            fc.ModifyAppliesToRange combinedRange
            On Error GoTo 0
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Bedingte Formatierung von A2 auf KW-Zellen erweitert."
    On Error GoTo 0
End Sub

'@Description("Finalizes a month by merging cells and adding borders")
Private Sub FinalizeMonth(ByVal targetSheet As Worksheet, _
                          ByVal dataRow As Long, _
                          ByVal startColumn As Long, _
                          ByVal endColumn As Long, _
                          ByVal monthName As String, _
                          ByVal yearValue As String)

    With targetSheet.Range(targetSheet.Cells(dataRow + MONTH_ROW_OFFSET, startColumn), _
                           targetSheet.Cells(dataRow + MONTH_ROW_OFFSET, endColumn))
        .Merge
        .value = monthName & " " & yearValue
        .NumberFormat = "MMMM YYYY"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 11
        '--- Only set outer borders, not inner vertical lines
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
    End With
End Sub

'@Description("Adds holidays and school vacations to the calendar")
Public Sub AddHolidaysAndVacations()
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet

    Dim datesRange As Range
    Set datesRange = targetSheet.Range("TAGE")

    Dim datesRowNumber As Long
    datesRowNumber = datesRange.Rows(1).Row

    '--- Process school vacations (Ferien)
    Dim vacationsTable As ListObject
    Set vacationsTable = Tabelle1.ListObjects("Ferien")

    Dim vacationRow As ListRow
    For Each vacationRow In vacationsTable.ListRows
        Call MarkVacationPeriod(targetSheet, vacationRow, datesRange, datesRowNumber)
    Next vacationRow

    '--- Process holidays (Feiertage)
    Dim holidaysTable As ListObject
    Set holidaysTable = Tabelle1.ListObjects("Feiertage")

    Dim holidayRow As ListRow
    For Each holidayRow In holidaysTable.ListRows
        Call MarkHoliday(targetSheet, holidayRow, datesRange, datesRowNumber)
    Next holidayRow

    'MsgBox "Feiertage und Schulferien wurden erfolgreich eingetragen.", vbInformation
    Application.StatusBar = False
End Sub

'@Description("Marks a vacation period in the calendar")
Private Sub MarkVacationPeriod(ByVal targetSheet As Worksheet, _
                                ByVal vacationRow As ListRow, _
                                ByVal datesRange As Range, _
                                ByVal datesRowNumber As Long)

    Dim vacationName As String
    Dim vacationStart As Date
    Dim vacationEnd As Date

    vacationName = vacationRow.Range.Cells(1, 1).value
    vacationStart = vacationRow.Range.Cells(1, 2).value
    vacationEnd = vacationRow.Range.Cells(1, 3).value

    Application.StatusBar = "Ferien / " & vacationName & " von " & vacationStart & " bis " & vacationEnd

    '--- Find start and end columns using .Find with multiple strategies
    Dim startCell As Range
    Dim endCell As Range

    '--- Strategy 1: Try finding with Date value directly
    Set startCell = datesRange.Find(What:=vacationStart, _
                                    LookIn:=xlValues, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)

    Set endCell = datesRange.Find(What:=vacationEnd, _
                                  LookIn:=xlValues, _
                                  LookAt:=xlWhole, _
                                  SearchOrder:=xlByColumns, _
                                  SearchDirection:=xlNext, _
                                  MatchCase:=False)

    '--- Strategy 2: If not found, try with CLng (integer date)
    If startCell Is Nothing Then
        Set startCell = datesRange.Find(What:=CLng(vacationStart), _
                                        LookIn:=xlValues, _
                                        LookAt:=xlWhole)
    End If

    If endCell Is Nothing Then
        Set endCell = datesRange.Find(What:=CLng(vacationEnd), _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole)
    End If

    '--- Strategy 3: If still not found, loop through manually
    If startCell Is Nothing Or endCell Is Nothing Then
        Dim currentCol As Long
        Dim cellValue As Variant

        For currentCol = datesRange.Column To datesRange.Column + datesRange.Columns.Count - 1
            cellValue = targetSheet.Cells(datesRowNumber, currentCol).value

            '--- Check if it's a date and matches
            If IsDate(cellValue) Then
                If CLng(CDate(cellValue)) = CLng(vacationStart) And startCell Is Nothing Then
                    Set startCell = targetSheet.Cells(datesRowNumber, currentCol)
                End If
                If CLng(CDate(cellValue)) = CLng(vacationEnd) And endCell Is Nothing Then
                    Set endCell = targetSheet.Cells(datesRowNumber, currentCol)
                End If
                If Not startCell Is Nothing And Not endCell Is Nothing Then Exit For
            End If
        Next currentCol
    End If

    '--- Mark vacation period if both dates found
    If Not startCell Is Nothing And Not endCell Is Nothing Then
        Dim firstColumn As Long
        Dim lastColumn As Long
        firstColumn = startCell.Column
        lastColumn = endCell.Column

        '--- FIX: Unmerge existing cells first to avoid conflicts
        Dim vacationRange As Range
        Set vacationRange = targetSheet.Range(targetSheet.Cells(datesRowNumber - 4, firstColumn), _
                                               targetSheet.Cells(datesRowNumber - 4, lastColumn))

        On Error Resume Next
        vacationRange.UnMerge
        On Error GoTo 0

        With vacationRange
            .Merge
            .value = vacationName
            .Font.Size = 6
            .HorizontalAlignment = xlCenter
            '--- Add background color
            .Interior.Pattern = xlSolid
            .Interior.Color = HOLIDAY_VACATION_COLOR
            '--- Only add borders if vacation name is not empty
            If Len(Trim$(vacationName)) > 0 Then
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(0, 0, 0)
                End With
            Else
                '--- Remove borders for empty vacation cells
                .Borders.LineStyle = xlNone
            End If
        End With
    Else
        Debug.Print "Ferien NICHT gefunden", vacationName, vacationStart, vacationEnd
    End If
End Sub

'@Description("Gets the actual date for a calendar column")
'@Param targetSheet The worksheet
'@Param dateRow The row containing date information
'@Param columnIndex The column to check
'@Returns The date for the specified column
Private Function GetDateForColumn(ByVal targetSheet As Worksheet, _
                                   ByVal dateRow As Long, _
                                   ByVal columnIndex As Long) As Date
    '--- Since we display weekday names, we need to calculate the actual date
    '--- We can look at the week number and date range in the header rows
    Dim weekNumber As Long
    Dim weekCell As Range
    
    '--- Find the calendar week for this column
    Set weekCell = targetSheet.Cells(dateRow + CALENDAR_WEEK_ROW_OFFSET, columnIndex).MergeArea.Resize(1, 1)
    
    If IsNumeric(weekCell.value) Then
        weekNumber = CLng(weekCell.value)
        
        '--- Calculate Monday of the ISO week using the helper function
        Dim baseDate As Date
        baseDate = GetMondayOfWeek(Year(Date), weekNumber)
        
        '--- Find which day of the week this column represents
        '--- by looking at the weekday name
        Dim weekdayName As String
        weekdayName = targetSheet.Cells(dateRow, columnIndex).value
        
        Select Case UCase(weekdayName)
            Case "MO", "MON"
                GetDateForColumn = baseDate
            Case "DI", "TUE"
                GetDateForColumn = baseDate + 1
            Case "MI", "WED"
                GetDateForColumn = baseDate + 2
            Case "DO", "THU"
                GetDateForColumn = baseDate + 3
            Case "FR", "FRI"
                GetDateForColumn = baseDate + 4
            Case Else
                GetDateForColumn = baseDate
        End Select
    Else
        '--- Fallback: return today's date
        GetDateForColumn = Date
    End If
End Function

'@Description("Marks a single holiday in the calendar")
Private Sub MarkHoliday(ByVal targetSheet As Worksheet, _
                        ByVal holidayRow As ListRow, _
                        ByVal datesRange As Range, _
                        ByVal datesRowNumber As Long)

    Dim holidayName As String
    Dim holidayDate As Date

    holidayName = holidayRow.Range.Cells(1, 1).value
    holidayDate = holidayRow.Range.Cells(1, 2).value

    Application.StatusBar = "Feiertag / " & holidayName & " " & holidayDate

    '--- Find date using .Find with multiple strategies
    Dim foundCell As Range

    '--- Strategy 1: Try finding with Date value directly
    Set foundCell = datesRange.Find(What:=holidayDate, _
                                    LookIn:=xlValues, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)

    '--- Strategy 2: If not found, try with CLng (integer date)
    If foundCell Is Nothing Then
        Set foundCell = datesRange.Find(What:=CLng(holidayDate), _
                                        LookIn:=xlValues, _
                                        LookAt:=xlWhole)
    End If

    '--- Strategy 3: If still not found, loop through manually
    If foundCell Is Nothing Then
        Dim currentCol As Long
        Dim cellValue As Variant

        For currentCol = datesRange.Column To datesRange.Column + datesRange.Columns.Count - 1
            cellValue = targetSheet.Cells(datesRowNumber, currentCol).value

            '--- Check if it's a date and matches
            If IsDate(cellValue) Then
                If CLng(CDate(cellValue)) = CLng(holidayDate) Then
                    Set foundCell = targetSheet.Cells(datesRowNumber, currentCol)
                    Exit For
                End If
            End If
        Next currentCol
    End If

    If Not foundCell Is Nothing Then
        Dim foundColumn As Long
        foundColumn = foundCell.Column

        '--- Color entire column
        With targetSheet.Range(targetSheet.Cells(datesRowNumber, foundColumn), _
                               targetSheet.Cells(datesRowNumber + EMPLOYEE_ROWS_COUNT, foundColumn)).Interior
            .Pattern = xlSolid
            .Color = HOLIDAY_VACATION_COLOR
        End With

        '--- Add holiday name
        With targetSheet.Cells(datesRowNumber - 8, foundColumn)
            .value = holidayName
            .Interior.Pattern = xlSolid
            .Interior.Color = HOLIDAY_VACATION_COLOR
        End With
    Else
        Debug.Print "Feiertag NICHT gefunden", holidayName, holidayDate
    End If
End Sub

'@Description("Applies conditional formatting to all tables")
'@Param useShortForm Whether to use short form codes (default: True)
'@Param startColumnIndex Starting column for formatting (default: 15)
Public Sub ApplyConditionalFormattingToTables(Optional ByVal useShortForm As Boolean = True, _
                                               Optional ByVal startColumnIndex As Long = 15)
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet

    Dim targetRange As Range
    Set targetRange = Nothing

    Dim listObj As ListObject
    Dim columnIndex As Long
    Dim tempRange As Range

    '--- Collect all cells from startColumnIndex+ in all tables
    For Each listObj In targetSheet.ListObjects
        If Not listObj.DataBodyRange Is Nothing Then
            For columnIndex = startColumnIndex To listObj.Range.Columns.Count
                If columnIndex <= listObj.ListColumns.Count Then
                    Set tempRange = listObj.ListColumns(columnIndex).DataBodyRange
                    If Not targetRange Is Nothing Then
                        Set targetRange = Union(targetRange, tempRange)
                    Else
                        Set targetRange = tempRange
                    End If
                End If
            Next columnIndex
        End If
    Next listObj

    If targetRange Is Nothing Then
        MsgBox "Keine gueltigen Zellen ab Spalte " & startColumnIndex & " gefunden.", vbExclamation
        Exit Sub
    End If

    '--- Clear existing conditional formatting
    targetRange.FormatConditions.Delete

    '--- Get absence codes
    Dim absenceCodes As Dictionary
    Set absenceCodes = AbsenceCode.GetAllCodes

    '--- Apply conditional formatting for each code
    Dim codeKey As Variant
    Dim currentAbsenceCode As AbsenceCode
    Dim formatFormula As String

    For Each codeKey In absenceCodes.Keys
        Set currentAbsenceCode = absenceCodes(codeKey)

        If useShortForm Then
            formatFormula = "=" & targetRange.Cells(1, 1).Address(False, False) & "=""" & currentAbsenceCode.ShortForm & """"
        Else
            formatFormula = "=" & targetRange.Cells(1, 1).Address(False, False) & "=""" & currentAbsenceCode.LongForm & """"
        End If

        With targetRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formatFormula)
            .StopIfTrue = False
            .Interior.Color = currentAbsenceCode.ColorRGB
        End With
    Next codeKey
End Sub

'@Description("Applies data validation dropdowns with absence codes to all tables")
'@Param startColumnIndex Starting column for validation (default: 15)
Public Sub ApplyDataValidationToTables(Optional ByVal startColumnIndex As Long = 15)
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet

    Dim targetRange As Range
    Set targetRange = Nothing

    Dim listObj As ListObject
    Dim columnIndex As Long
    Dim tempRange As Range

    '--- Collect all cells from startColumnIndex+ in all tables
    For Each listObj In targetSheet.ListObjects
        If Not listObj.DataBodyRange Is Nothing Then
            For columnIndex = startColumnIndex To listObj.Range.Columns.Count
                If columnIndex <= listObj.ListColumns.Count Then
                    Set tempRange = listObj.ListColumns(columnIndex).DataBodyRange
                    If Not targetRange Is Nothing Then
                        Set targetRange = Union(targetRange, tempRange)
                    Else
                        Set targetRange = tempRange
                    End If
                End If
            Next columnIndex
        End If
    Next listObj

    If targetRange Is Nothing Then
        MsgBox "Keine gueltigen Zellen ab Spalte " & startColumnIndex & " gefunden.", vbExclamation
        Exit Sub
    End If
    
    '--- Change formating of cells in the calendar
    With targetRange
        .Font.Size = 6
        .Font.Name = "Arial"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    
    '--- Clear existing validation
    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo 0

    '--- Get absence codes for dropdown
    Dim absenceCodes As Dictionary
    Set absenceCodes = AbsenceCode.GetAllCodes

    '--- Build comma-separated list of short form codes
    Dim validationList As String
    validationList = ""

    Dim codeKey As Variant
    Dim currentAbsenceCode As AbsenceCode

    For Each codeKey In absenceCodes.Keys
        Set currentAbsenceCode = absenceCodes(codeKey)
        If Len(validationList) > 0 Then
            validationList = validationList & ","
        End If
        validationList = validationList & currentAbsenceCode.ShortForm
    Next codeKey

    '--- Apply data validation
    If Len(validationList) > 0 Then
        With targetRange.Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:=validationList
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = False
        End With
    End If
End Sub


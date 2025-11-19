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

'@Description("Creates a calendar with work days only (Monday-Friday)")
'@Param startCell The cell where the calendar should start
Public Sub CreateWorkDayCalendar(ByVal startCell As Range)
    If startCell Is Nothing Then
        MsgBox "Keine Startzelle ausgewählt.", vbExclamation
        Exit Sub
    End If

    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet

    '--- Get date range from user
    Dim startDate As Date
    Dim endDate As Date

    startDate = Application.InputBox("Startdatum eingeben (z.B. 01.01.2025):", "Startdatum", Date, , , , , 1)
    endDate = Application.InputBox("Enddatum eingeben (z.B. 31.12.2025):", "Enddatum", Date + 30, , , , , 1)

    If endDate < startDate Then
        MsgBox "Enddatum muss nach dem Startdatum liegen!", vbExclamation
        Exit Sub
    End If

    '--- Performance optimization
    Dim originalScreenUpdating As Boolean
    Dim originalCursor As XlMousePointer

    originalScreenUpdating = Application.ScreenUpdating
    originalCursor = Application.Cursor

    Application.ScreenUpdating = False
    Application.Cursor = xlWait

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
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With

    '--- Main loop: iterate through all dates
    Do While currentDate <= endDate
        '--- Only process weekdays (Monday-Friday)
        If Weekday(currentDate, vbMonday) <= 5 Then
            Application.StatusBar = currentYear & " / " & currentMonth & " / " & currentCalendarWeek & " / " & currentDate

            '--- Write date
            targetSheet.Cells(currentRow, currentColumn).value = currentDate
            targetSheet.Cells(currentRow, currentColumn).NumberFormat = "dd"
            targetSheet.Cells(currentRow, currentColumn).HorizontalAlignment = xlCenter
            targetSheet.Columns(currentColumn).ColumnWidth = 0.69

            '--- Format holiday/vacation cell (merged vertical)
            With targetSheet.Range(targetSheet.Cells(currentRow - 5, currentColumn), _
                                   targetSheet.Cells(currentRow - 8, currentColumn))
                .Merge
                .Font.Size = 6
                .Orientation = 90
                .VerticalAlignment = xlBottom
                .HorizontalAlignment = xlCenter
            End With

            '--- Check for calendar week change
            If WorksheetFunction.WeekNum(currentDate, 2) <> currentCalendarWeek Then
                Call FinalizeCalendarWeek(targetSheet, currentRow, weekStartColumn, currentColumn - 1, currentCalendarWeek)

                '--- Draw border for new week
                With targetSheet.Range(targetSheet.Cells(currentRow, currentColumn), _
                                       targetSheet.Cells(currentRow + EMPLOYEE_ROWS_COUNT, currentColumn)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
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

    Application.ScreenUpdating = originalScreenUpdating
    Application.Cursor = originalCursor
    Application.StatusBar = False

    MsgBox "Kalender mit Arbeitstagen erfolgreich erstellt!", vbInformation

    '--- Ask if holidays should be added
    Dim addHolidays As VbMsgBoxResult
    addHolidays = MsgBox("Sollen die Feiertage auch eingetragen werden?", vbYesNo, "Feiertage eintragen")

    If addHolidays = vbYes Then
        Call AddHolidaysAndVacations
    End If

    Call ApplyConditionalFormattingToTables

    Tabelle1.Activate
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
        .Font.Bold = True
        .Font.Size = 10
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
    End With

    '--- Date range (e.g., "01-05")
    Dim firstDayDate As Date
    Dim lastDayDate As Date
    firstDayDate = targetSheet.Cells(dataRow, startColumn).value
    lastDayDate = targetSheet.Cells(dataRow, endColumn).value

    With targetSheet.Range(targetSheet.Cells(dataRow + DATE_ROW_OFFSET, startColumn), _
                           targetSheet.Cells(dataRow + DATE_ROW_OFFSET, endColumn))
        .Merge
        .value = Format(firstDayDate, "dd") & "-" & Format(lastDayDate, "dd")
        .HorizontalAlignment = xlCenter
        .Font.Bold = False
        .Font.Size = 8
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End With
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
        .Font.Bold = True
        .Font.Size = 11
        With .Borders
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

    MsgBox "Feiertage und Schulferien wurden erfolgreich eingetragen.", vbInformation
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

    Dim firstColumn As Long
    Dim lastColumn As Long
    firstColumn = 0
    lastColumn = 0

    '--- Find columns for vacation period
    Dim dateCell As Range
    For Each dateCell In datesRange.Cells
        If IsDate(dateCell.value) Then
            If dateCell.value >= vacationStart And dateCell.value <= vacationEnd Then
                If firstColumn = 0 Then firstColumn = dateCell.Column
                lastColumn = dateCell.Column
            End If
        End If
    Next dateCell

    '--- Mark vacation period
    If firstColumn > 0 And lastColumn >= firstColumn Then
        With targetSheet.Range(targetSheet.Cells(datesRowNumber - 4, firstColumn), _
                               targetSheet.Cells(datesRowNumber - 4, lastColumn))
            .Merge
            .value = vacationName
            .Font.Size = 6
            .HorizontalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .Color = RGB(0, 0, 0)
            End With
        End With
    End If
End Sub

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

    '--- Find date column
    Dim dateCell As Range
    Set dateCell = Nothing

    Dim currentCell As Range
    For Each currentCell In datesRange.Cells
        If IsDate(currentCell.value) Then
            If CLng(currentCell.value) = CLng(holidayDate) Then
                Set dateCell = currentCell
                Exit For
            End If
        End If
    Next currentCell

    If Not dateCell Is Nothing Then
        '--- Color entire column
        With targetSheet.Range(targetSheet.Cells(datesRowNumber, dateCell.Column), _
                               targetSheet.Cells(datesRowNumber + EMPLOYEE_ROWS_COUNT, dateCell.Column)).Interior
            .Pattern = xlSolid
            .ColorIndex = 33
        End With

        '--- Add holiday name
        With targetSheet.Cells(datesRowNumber - 8, dateCell.Column)
            .value = holidayName
            .Interior.Pattern = xlSolid
            .Interior.ColorIndex = 33
        End With
    Else
        Debug.Print "Feiertag NICHT gefunden", holidayName, holidayDate
    End If
End Sub

'@Description("Applies conditional formatting and dropdowns to all tables")
Public Sub ApplyConditionalFormattingToTables(Optional ByVal useShortForm As Boolean = True, _
                                               Optional ByVal startColumnIndex As Long = 15)
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet

    Dim targetRange As Range
    Set targetRange = Nothing

    Dim listObj As ListObject
    Dim columnIndex As Long
    Dim tempRange As Range

    '--- Collect all cells from column 15+ in all tables
    For Each listObj In targetSheet.ListObjects
        For columnIndex = startColumnIndex To listObj.Range.Columns.Count
            Set tempRange = listObj.ListColumns(columnIndex).DataBodyRange
            If Not targetRange Is Nothing Then
                Set targetRange = Union(targetRange, tempRange)
            Else
                Set targetRange = tempRange
            End If
        Next columnIndex
    Next listObj

    If targetRange Is Nothing Then
        MsgBox "Keine gültigen Zellen ab Spalte " & startColumnIndex & " gefunden.", vbExclamation
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

    '--- Clear existing validation
    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo 0
End Sub

Attribute VB_Name = "WeeklySheetService"
'@Folder("Services.WeeklySheet")
'@ModuleDescription("Creates and manages calendar week (KW) sheets from main planner")
Option Explicit

'@Description("Creates a new KW sheet from template, copying relevant employee data")
'@Param selectedCell Cell containing the calendar week number
Public Sub CreateWeeklySheet(ByVal selectedCell As Range)
    Application.ScreenUpdating = False
    Application.StatusBar = "Wochenplan erstellen ..."
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    '--- Parse calendar week from cell
    Dim calendarWeekNumber As Long
    On Error Resume Next
    calendarWeekNumber = CLng(selectedCell.offset(0, 0).value)
    On Error GoTo 0

    If calendarWeekNumber = 0 Then
        MsgBox "Keine gültige Kalenderwoche ausgewählt!", vbExclamation
        Exit Sub
    End If

    '--- Get week dates
    Dim weekStartDate As Date
    Dim weekEndDate As Date
    weekStartDate = selectedCell.offset(2, 0).value
    weekEndDate = selectedCell.offset(2, 4).value

    Dim sheetName As String
    sheetName = "KW" & calendarWeekNumber & " " & Format(weekStartDate, "YYYY")

    '--- Check if sheet already exists
    Dim existingSheet As Worksheet
    For Each existingSheet In ThisWorkbook.Worksheets
        If existingSheet.Name = sheetName Then
            existingSheet.Visible = xlSheetVisible
            existingSheet.Activate
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Application.DisplayAlerts = True
            Exit Sub
        End If
    Next existingSheet

    '--- Copy template
    Dim newSheet As Worksheet
    With Tabelle7
        .copying = True
        .Visible = xlSheetVisible
        .Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set newSheet = ActiveSheet

        On Error Resume Next
        newSheet.Name = sheetName
        On Error GoTo 0

        .Visible = xlSheetHidden
        .copying = False
    End With

    '--- Fill sheet header
    newSheet.Range("A3:A4").value = "KW" & calendarWeekNumber
    newSheet.Range("E4").value = weekStartDate
    newSheet.Range("F4").value = weekEndDate
    newSheet.Range("J3").value = Now()

    Application.StatusBar = "Das Blatt '" & sheetName & "' wurde neu erstellt."

    '--- Copy employee data
    Call CopyEmployeeDataToWeeklySheet(newSheet, selectedCell)

    '--- Replace short codes with long form
    Call ReplaceAbsenceCodesWithLongForm(newSheet)

    '--- Apply formatting
    CalendarService.ApplyConditionalFormattingToTables useShortForm:=False, startColumnIndex:=5

    '--- Initialize filter listboxes
    Call InitializeFilterListBox(newSheet, "ListBoxFunktion", "Funktion")
    Call InitializeFilterListBox(newSheet, "ListBoxTeam", "Team")

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculate
    Application.StatusBar = "DONE!"
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

'@Description("Copies employee data from main planner to weekly sheet")
Private Sub CopyEmployeeDataToWeeklySheet(ByVal weeklySheet As Worksheet, ByVal selectedCell As Range)
    Dim startColumn As Long
    Dim endColumn As Long
    startColumn = selectedCell.Column
    endColumn = startColumn + 4 '5 weekdays

    Dim sourceTable As ListObject
    Dim sourceRow As ListRow
    Dim weeklyTable As ListObject
    Dim newRow As ListRow

    Set weeklyTable = weeklySheet.Range("A7").ListObject

    Application.ScreenUpdating = False

    '--- Copy from all tables in Tabelle3
    For Each sourceTable In Tabelle3.ListObjects
        For Each sourceRow In sourceTable.ListRows
            '--- Skip rows without name
            If sourceRow.Range(1, 7).value = vbNullString Then GoTo SkipRow

            '--- Add new row to weekly sheet
            Set newRow = weeklyTable.ListRows.Add

            '--- Copy employee data
            newRow.Range(1, 1).value = sourceRow.Range(1, 6).value 'Number
            newRow.Range(1, 2).value = sourceRow.Range(1, 7).value & vbNewLine & _
                                       sourceRow.Range(1, 9).value & vbNewLine & _
                                       sourceRow.Range(1, 13).value 'Name, Phone, Email

            DateHelpers.FormatFirstLineBold newRow.Range(1, 2)

            newRow.Range(1, 3).value = sourceRow.Range(1, 8).value 'Function
            newRow.Range(1, 4).value = sourceRow.Range(1, 10).value 'Team

            '--- Copy weekday data (5 columns)
            Dim dayIndex As Long
            Dim targetColumn As Long
            targetColumn = 5

            For dayIndex = startColumn To endColumn
                newRow.Range(1, targetColumn).value = sourceRow.Range(1, dayIndex).value
                targetColumn = targetColumn + 1
            Next dayIndex

SkipRow:
        Next sourceRow
    Next sourceTable
End Sub

'@Description("Replaces short absence codes with long form text")
Private Sub ReplaceAbsenceCodesWithLongForm(ByVal weeklySheet As Worksheet)
    Dim weeklyTable As ListObject
    Set weeklyTable = weeklySheet.Range("A7").ListObject

    Dim absenceCodes As Dictionary
    Set absenceCodes = AbsenceCode.GetAllCodes

    Dim currentCell As Range
    Dim codeKey As Variant
    Dim currentAbsenceCode As AbsenceCode

    For Each currentCell In weeklyTable.DataBodyRange.Cells
        For Each codeKey In absenceCodes.Keys
            Set currentAbsenceCode = absenceCodes(codeKey)

            If currentCell.value = currentAbsenceCode.ShortForm Then
                currentCell.value = currentAbsenceCode.LongForm
                Exit For
            End If
        Next codeKey
    Next currentCell
End Sub

'@Description("Initializes ActiveX ListBox with unique values from table column")
'@Param targetSheet The worksheet containing the ListBox
'@Param listBoxName Name of the ActiveX ListBox control
'@Param columnName Name of the table column to read from
Public Sub InitializeFilterListBox( _
        ByVal targetSheet As Worksheet, _
        ByVal listBoxName As String, _
        ByVal columnName As String)

    Application.StatusBar = "Initialisiere " & listBoxName

    Dim dataTable As ListObject
    Set dataTable = targetSheet.ListObjects(1)

    Dim listBox As MSForms.ListBox
    On Error Resume Next
    Set listBox = targetSheet.OLEObjects(listBoxName).Object
    On Error GoTo 0

    If listBox Is Nothing Then
        MsgBox "ListBox '" & listBoxName & "' nicht gefunden auf Blatt " & targetSheet.Name, vbExclamation
        Exit Sub
    End If

    Dim columnRange As Range
    On Error Resume Next
    Set columnRange = dataTable.ListColumns(columnName).DataBodyRange
    On Error GoTo 0

    If columnRange Is Nothing Then
        listBox.Clear
        Exit Sub
    End If

    '--- Collect unique values
    Dim uniqueValues As Dictionary
    Set uniqueValues = EmployeeService.GetUniqueValuesFromRange(columnRange, extractFirstLineOnly:=True)

    '--- Populate ListBox
    listBox.Clear

    Dim uniqueKey As Variant
    For Each uniqueKey In uniqueValues.Keys
        listBox.AddItem CStr(uniqueKey)
    Next uniqueKey
End Sub

Attribute VB_Name = "WeeklyReportService"
'@Folder("Services.WeeklyReport")
'@ModuleDescription("Creates and manages weekly reports (Wochenrapporte) for employees")
Option Explicit

Private Const REPORT_FERIEN_ROW As Long = 26
Private Const REPORT_MILITAR_ROW As Long = 27
Private Const REPORT_UNFALL_ROW As Long = 28
Private Const REPORT_KRANK_ROW As Long = 29
Private Const REPORT_PROJECT_COLUMN As Long = 14 'Column N

'@Description("Creates weekly reports for all employees in the active week sheet")
Public Sub CreateWeeklyReports()
    Dim weeklySheet As Worksheet
    Set weeklySheet = ActiveSheet

    Dim calendarWeek As String
    Dim weekStartDate As Date
    Dim weekEndDate As Date

    calendarWeek = weeklySheet.Range("A3").value
    weekStartDate = weeklySheet.Range("E4").value
    weekEndDate = weeklySheet.Range("F4").value

    '--- Performance optimization
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEnableEvents As Boolean
    Dim originalDisplayAlerts As Boolean

    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalEnableEvents = Application.EnableEvents
    originalDisplayAlerts = Application.DisplayAlerts

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    '--- Collect unique projects from columns E:I
    Dim weekTable As ListObject
    Set weekTable = weeklySheet.Range("E7").ListObject

    Dim projectRange As Range
    Set projectRange = Intersect(weekTable.DataBodyRange, weeklySheet.Range("E:I"))

    Dim uniqueProjects As Dictionary
    Set uniqueProjects = EmployeeService.GetUniqueValuesFromRange(projectRange, includeHidden:=False, extractFirstLineOnly:=True)

    '--- Filter out known projects from settings
    Call RemoveKnownProjects(uniqueProjects)

    '--- Prompt for new project details
    Dim projectKey As Variant
    Dim projectObj As Project

    For Each projectKey In uniqueProjects.Keys
        Set projectObj = ProjectService.PromptForProjectDetails(CStr(projectKey))

        If projectObj Is Nothing Then
            '--- User cancelled
            GoTo CleanupAndExit
        End If

        '--- Store project in dictionary for later use
        uniqueProjects(projectKey) = projectObj.ToStorageString
    Next projectKey

    '--- Collect unique employees from column A
    Dim employeeRange As Range
    Set employeeRange = Intersect(weekTable.DataBodyRange, weeklySheet.Range("A:A"))

    Dim uniqueEmployees As Dictionary
    Set uniqueEmployees = EmployeeService.GetUniqueValuesFromRange(employeeRange, includeHidden:=False)

    '--- Create new workbook for reports
    Dim workbookPath As String
    workbookPath = ActiveWorkbook.Path

    Dim reportsWorkbook As Workbook
    Set reportsWorkbook = Workbooks.Add
    reportsWorkbook.SaveAs workbookPath & "\Wochenrapporte_" & calendarWeek & ".xlsm", xlOpenXMLWorkbookMacroEnabled

    With reportsWorkbook.Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
        .EnableEvents = False
    End With

    Dim createdReportsCount As Long
    createdReportsCount = 0

    '--- Create report for each employee
    Dim employeeKey As Variant
    Dim employeeCell As Range
    Dim employeeRow As Long
    Dim employeeName As String

    For Each employeeKey In uniqueEmployees.Keys
        On Error Resume Next

        Set employeeCell = employeeRange.Find(employeeKey)
        If employeeCell Is Nothing Then GoTo NextEmployee

        employeeRow = employeeCell.Row

        '--- Check skip flag (column K)
        If Not weeklySheet.Cells(employeeRow, 11).value Then
            employeeName = Split(weeklySheet.Cells(employeeRow, 2).value, vbLf)(0)

            Call CreateSingleEmployeeReport( _
                reportsWorkbook, _
                weeklySheet, _
                employeeRow, _
                employeeName, _
                calendarWeek, _
                weekStartDate, _
                weekEndDate, _
                uniqueProjects)

            createdReportsCount = createdReportsCount + 1
        End If

NextEmployee:
        On Error GoTo ErrorHandler
    Next employeeKey

CleanupAndExit:
    '--- Delete empty first sheet
    If reportsWorkbook.Sheets.Count > 1 Then
        Application.DisplayAlerts = False
        reportsWorkbook.Sheets(1).Delete
        Application.DisplayAlerts = True
    End If

    '--- Calculate and restore settings
    reportsWorkbook.Application.Calculate

    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = True

    shWRTemplate.Visible = xlSheetHidden

    MsgBox createdReportsCount & " Wochenrapporte wurden erfolgreich erstellt!", vbInformation, "Rapporte erstellt"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents

    MsgBox "Fehler " & Err.Number & " ist aufgetreten:" & vbNewLine & vbNewLine & _
           Err.Description & vbNewLine & vbNewLine & _
           "Quelle: " & Err.Source, _
           vbCritical, "Fehler bei der Rapport-Erstellung"

    Debug.Print "Fehler in CreateWeeklyReports: " & Err.Number & " - " & Err.Description
End Sub

'@Description("Removes already known projects from the unique projects dictionary")
Private Sub RemoveKnownProjects(ByRef uniqueProjects As Dictionary)
    '--- FIX: Use ProjectService.GetProjectSheet() instead of non-existent Tabelle5
    '--- The original code referenced Tabelle5.ListObjects("Tabelle6") which caused
    '--- "Blatt 'Projektnummern' nicht gefunden" error because Tabelle5.doccls doesn't exist
    Dim projectSheet As Worksheet
    Set projectSheet = ProjectService.GetProjectSheet()

    If projectSheet Is Nothing Then
        Exit Sub '--- No project sheet found, skip removing known projects
    End If

    '--- Try to find the project ListObject (could be first table on sheet)
    Dim knownProjectsRange As Range
    On Error Resume Next
    If projectSheet.ListObjects.Count > 0 Then
        '--- Use the first ListObject's second column (project names)
        Set knownProjectsRange = projectSheet.ListObjects(1).DataBodyRange.Columns(2)
    End If
    On Error GoTo 0

    '--- Fallback: If no ListObject, use UsedRange approach
    If knownProjectsRange Is Nothing Then
        Dim lastRow As Long
        lastRow = projectSheet.Cells(projectSheet.Rows.Count, 1).End(xlUp).Row
        If lastRow > 1 Then
            Set knownProjectsRange = projectSheet.Range(projectSheet.Cells(2, 2), projectSheet.Cells(lastRow, 2))
        Else
            Exit Sub '--- No data
        End If
    End If

    Dim projectCell As Range
    Dim projectLines() As String
    Dim firstLine As String

    For Each projectCell In knownProjectsRange.Cells
        If Not IsEmpty(projectCell.value) And Not IsError(projectCell.value) Then
            projectLines = Split(projectCell.value, Chr(10))

            If UBound(projectLines) > 0 Then
                firstLine = projectLines(0)
            Else
                firstLine = projectCell.Value2
            End If

            If uniqueProjects.Exists(firstLine) Then
                uniqueProjects.Remove firstLine
            End If
        End If
    Next projectCell
End Sub

'@Description("Creates a single employee's weekly report")
Private Sub CreateSingleEmployeeReport( _
        ByVal reportsWorkbook As Workbook, _
        ByVal weeklySheet As Worksheet, _
        ByVal employeeRow As Long, _
        ByVal employeeName As String, _
        ByVal calendarWeek As String, _
        ByVal weekStartDate As Date, _
        ByVal weekEndDate As Date, _
        ByVal projectsDict As Dictionary)

    '--- Copy template
    shWRTemplate.Visible = xlSheetVisible
    shWRTemplate.Copy After:=reportsWorkbook.Sheets(reportsWorkbook.Sheets.Count)

    Dim reportSheet As Worksheet
    Set reportSheet = reportsWorkbook.Sheets(reportsWorkbook.Sheets.Count)
    reportSheet.Name = employeeName

    '--- Fill header
    reportSheet.Range("A2").value = "Wochenrapport von: " & employeeName
    reportSheet.Range("E2").value = "Datum von: " & Format(weekStartDate, "DD.MM.YYYY")
    reportSheet.Range("J2").value = "bis: " & Format(weekEndDate, "DD.MM.YYYY")
    reportSheet.Range("N2").value = "Kalenderwoche: " & Right(calendarWeek, 2)

    '--- Process each weekday (columns E:I = Mon-Fri)
    Dim dayIndex As Long
    Dim dayCell As Range

    dayIndex = 0
    For Each dayCell In weeklySheet.Range(weeklySheet.Cells(employeeRow, 5), _
                                          weeklySheet.Cells(employeeRow, 9)).Cells
        dayIndex = dayIndex + 1

        Call ProcessDayEntry(reportSheet, dayCell, dayIndex, projectsDict)
    Next dayCell
End Sub

'@Description("Processes a single day's entry for an employee")
Private Sub ProcessDayEntry( _
        ByVal reportSheet As Worksheet, _
        ByVal dayCell As Range, _
        ByVal dayIndex As Long, _
        ByVal projectsDict As Dictionary)

    Dim cellValue As String
    cellValue = dayCell.value

    Dim projectRow As Long
    Dim hoursColumnIndex As Long
    hoursColumnIndex = dayIndex + 2 '(Monday=1 -> Column C=3)

    Select Case cellValue
        Case "Krank"
            projectRow = REPORT_KRANK_ROW
            reportSheet.Cells(projectRow, hoursColumnIndex).value = 8

        Case "Unfall"
            projectRow = REPORT_UNFALL_ROW
            reportSheet.Cells(projectRow, hoursColumnIndex).value = 8

        Case "Militär"
            projectRow = REPORT_MILITAR_ROW
            reportSheet.Cells(projectRow, hoursColumnIndex).value = 8

        Case "Ferien"
            projectRow = REPORT_FERIEN_ROW
            reportSheet.Cells(projectRow, hoursColumnIndex).value = 8

        Case "Schule", "Überbetr.Kurs"
            '--- Not tracked in report

        Case ""
            '--- Empty cell

        Case Else
            '--- Regular project
            Call AddProjectHours(reportSheet, dayCell, dayIndex, projectsDict)
    End Select
End Sub

'@Description("Adds project hours to the report, handling comments")
Private Sub AddProjectHours( _
        ByVal reportSheet As Worksheet, _
        ByVal dayCell As Range, _
        ByVal dayIndex As Long, _
        ByVal projectsDict As Dictionary)

    Dim cellValue As String
    cellValue = dayCell.value

    Dim projectName As String
    Dim projectComment As String
    Dim cellLines() As String

    '--- Extract project name and comment
    If InStr(cellValue, Chr(10)) > 0 Then
        cellLines = Split(cellValue, Chr(10))
        projectName = cellLines(0)
        If UBound(cellLines) >= 1 Then
            projectComment = cellLines(1)
        End If
    Else
        projectName = cellValue
        projectComment = vbNullString
    End If

    '--- Find or create project row
    Dim projectColumn As Range
    Set projectColumn = reportSheet.UsedRange.Resize(, 1).offset(0, 13) 'Column N

    Dim projectRow As Long
    Dim existingProjectCell As Range
    Set existingProjectCell = projectColumn.Find(projectName)

    If Not existingProjectCell Is Nothing Then
        '--- Project exists
        projectRow = existingProjectCell.Row
    Else
        '--- New project - add row
        projectRow = reportSheet.Cells(24, 1).End(xlUp).Row + 1

        reportSheet.Cells(projectRow, REPORT_PROJECT_COLUMN).value = projectName

        '--- Get commission and remarks from dictionary
        If projectsDict.Exists(projectName) Then
            Dim projectData() As String
            projectData = Split(projectsDict(projectName), ";")

            If UBound(projectData) > 0 Then
                reportSheet.Cells(projectRow, 1).value = projectData(1) 'Remarks
                reportSheet.Cells(projectRow, 2).value = projectData(0) 'Commission
            Else
                reportSheet.Cells(projectRow, 1).value = projectsDict(projectName)
            End If
        End If
    End If

    '--- Add hours and comment
    Dim hoursColumnIndex As Long
    hoursColumnIndex = dayIndex + 2

    With reportSheet.Cells(projectRow, hoursColumnIndex)
        .value = 8.5

        If Len(projectComment) > 0 Then
            .AddComment
            .Comment.Text projectComment
        End If
    End With
End Sub

'@Description("Sends email reminder to all employees to submit weekly reports")
Public Sub SendWeeklyReportReminder()
    Dim weeklySheet As Worksheet
    Set weeklySheet = ActiveSheet

    Dim calendarWeek As String
    calendarWeek = weeklySheet.Range("A3").value

    '--- Performance optimization
    Dim originalScreenUpdating As Boolean
    originalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    '--- Collect employees
    Dim weekTable As ListObject
    Set weekTable = weeklySheet.Range("E7").ListObject

    Dim employeeRange As Range
    Set employeeRange = Intersect(weekTable.DataBodyRange, weeklySheet.Range("A:A"))

    Dim uniqueEmployees As Dictionary
    Set uniqueEmployees = EmployeeService.GetUniqueValuesFromRange(employeeRange, includeHidden:=False)

    '--- Build email recipient list
    Dim emailList As String
    Dim emailCount As Long
    Dim errorCount As Long
    Dim errorList As String

    Dim employeeKey As Variant
    Dim employeeCell As Range
    Dim employeeRow As Long
    Dim emp As Employee

    For Each employeeKey In uniqueEmployees.Keys
        On Error Resume Next

        Set employeeCell = employeeRange.Find(employeeKey)
        If Not employeeCell Is Nothing Then
            employeeRow = employeeCell.Row

            '--- Check skip flag
            If Not weeklySheet.Cells(employeeRow, 11).value Then
                Set emp = New Employee
                emp.ParseFromCellValue weeklySheet.Cells(employeeRow, 2).value

                If emp.HasValidEmail Then
                    If emailList <> vbNullString Then emailList = emailList & "; "
                    emailList = emailList & emp.EmailAddress
                    emailCount = emailCount + 1
                Else
                    errorCount = errorCount + 1
                    errorList = errorList & "- " & emp.DisplayName & ": Keine gültige E-Mail-Adresse" & vbNewLine
                End If
            End If
        End If

        On Error GoTo ErrorHandler
    Next employeeKey

    If emailList = vbNullString Then
        MsgBox "Keine gültigen E-Mail-Adressen gefunden." & vbNewLine & vbNewLine & _
               "Bitte prüfen Sie, ob die E-Mail-Adressen in der 3. Zeile der Namenszellen stehen.", _
               vbExclamation, "Keine Empfänger"
        GoTo CleanupAndExit
    End If

    '--- Create Outlook email
    Dim outlookApp As Object
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo ErrorHandler

    If outlookApp Is Nothing Then
        MsgBox "Outlook konnte nicht gestartet werden.", vbCritical, "Outlook-Fehler"
        GoTo CleanupAndExit
    End If

    Dim mailItem As Object
    Set mailItem = outlookApp.CreateItem(0) 'olMailItem

    '--- FIX: Use HTMLBody with UTF-8 charset to properly display German umlauts
    '--- The original .body property doesn't properly encode UTF-8 characters
    '--- causing umlauts (ä, ö, ü) to display as wrong characters
    Dim emailBodyHTML As String
    emailBodyHTML = "<!DOCTYPE html>" & vbNewLine & _
                    "<html>" & vbNewLine & _
                    "<head>" & vbNewLine & _
                    "<meta charset=""UTF-8"">" & vbNewLine & _
                    "</head>" & vbNewLine & _
                    "<body style=""font-family: Calibri, Arial, sans-serif; font-size: 11pt;"">" & vbNewLine & _
                    "<p>Hallo zusammen,</p>" & vbNewLine & _
                    "<p>bitte gebt noch euren Wochenrapport ab.</p>" & vbNewLine & _
                    "<p>Vielen Dank!<br>" & vbNewLine & _
                    "Mit freundlichen Grüssen</p>" & vbNewLine & _
                    "</body>" & vbNewLine & _
                    "</html>"

    With mailItem
        .To = emailList
        .Subject = "Erinnerung: Wochenrapport " & calendarWeek & " abgeben"
        .HTMLBody = emailBodyHTML
        .Importance = 2 'olImportanceHigh
        .Display 'Show email for review
    End With

CleanupAndExit:
    Application.ScreenUpdating = originalScreenUpdating

    If emailCount > 0 Then
        Dim successMessage As String
        successMessage = "E-Mail an " & emailCount & " Empfänger wurde erstellt."

        If errorCount > 0 Then
            successMessage = successMessage & vbNewLine & vbNewLine & _
                            errorCount & " Fehler aufgetreten:" & vbNewLine & vbNewLine & _
                            errorList
            MsgBox successMessage, vbExclamation, "E-Mail erstellt mit Fehlern"
        Else
            MsgBox successMessage, vbInformation, "E-Mail erfolgreich erstellt"
        End If
    End If

    Set mailItem = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = originalScreenUpdating

    MsgBox "Fehler " & Err.Number & " ist aufgetreten:" & vbNewLine & vbNewLine & _
           Err.Description & vbNewLine & vbNewLine & _
           "Quelle: " & Err.Source, _
           vbCritical, "Fehler beim E-Mail-Versand"

    Debug.Print "Fehler in SendWeeklyReportReminder: " & Err.Number & " - " & Err.Description

    Set mailItem = Nothing
    Set outlookApp = Nothing
End Sub

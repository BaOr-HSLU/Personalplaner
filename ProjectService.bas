Attribute VB_Name = "ProjectService"
'@Folder("Services.Project")
'@ModuleDescription("Service for managing project data, commission numbers, and remarks")
Option Explicit

Private Const PROJECT_SHEET_NAME As String = "Projektnummern"

'@Description("Loads a project from the project master sheet")
'@Param projectName The name of the project to load
'@Returns Project object or Nothing if not found
Public Function LoadProject(ByVal projectName As String) As Project
    '@Ignore EmptyStringLiteral
    On Error GoTo ErrorHandler

    Dim projectSheet As Worksheet
    Set projectSheet = GetProjectSheet()

    If projectSheet Is Nothing Then
        Set LoadProject = Nothing
        Exit Function
    End If

    '--- Search for project in column A
    Dim foundCell As Range
    Set foundCell = projectSheet.UsedRange.Resize(, 1).Find(projectName)

    If foundCell Is Nothing Then
        Set LoadProject = Nothing
        Exit Function
    End If

    '--- Create and populate project object
    Dim proj As Project
    Set proj = New Project

    proj.projectName = projectName
    proj.CommissionNumber = projectSheet.Cells(foundCell.Row, 2).value
    proj.Remarks = projectSheet.Cells(foundCell.Row, 3).value

    Set LoadProject = proj
    Exit Function

ErrorHandler:
    Set LoadProject = Nothing
End Function

'@Description("Saves a project to the project master sheet")
'@Param proj The project to save
'@Returns True if successful
Public Function SaveProject(ByVal proj As Project) As Boolean
    On Error GoTo ErrorHandler

    Dim projectSheet As Worksheet
    Set projectSheet = GetProjectSheet()

    If projectSheet Is Nothing Then
        SaveProject = False
        Exit Function
    End If

    If Not proj.IsValid Then
        SaveProject = False
        Exit Function
    End If

    '--- Check if project already exists
    Dim foundCell As Range
    Set foundCell = projectSheet.UsedRange.Resize(, 1).Find(proj.projectName)

    Dim targetRow As Long

    If foundCell Is Nothing Then
        '--- New project: add at end
        targetRow = projectSheet.UsedRange.Rows.Count + 1
    Else
        '--- Existing project: update
        targetRow = foundCell.Row
    End If

    '--- Write project data
    projectSheet.Cells(targetRow, 1).value = proj.projectName
    projectSheet.Cells(targetRow, 2).value = proj.CommissionNumber
    projectSheet.Cells(targetRow, 3).value = proj.Remarks

    SaveProject = True
    Exit Function

ErrorHandler:
    SaveProject = False
End Function

'@Description("Prompts user to enter commission number and remarks for a project")
'@Param projectName The project name to get details for
'@Returns Project object with user input, or Nothing if cancelled
Public Function PromptForProjectDetails(ByVal projectName As String) As Project
    '@Ignore EmptyStringLiteral
    Dim proj As Project
    Set proj = New Project
    proj.projectName = projectName

    '--- Check if project exists, offer to load
    Dim existingProject As Project
    Set existingProject = LoadProject(projectName)

    If Not existingProject Is Nothing Then
        '--- Ask if user wants to use existing data
        Dim useExisting As VbMsgBoxResult
        useExisting = MsgBox( _
            "Sollen die Projektdaten geladen werden?" & vbNewLine & vbNewLine & _
            "Projekt: " & existingProject.projectName & vbNewLine & _
            "Kommission: " & existingProject.CommissionNumber & vbNewLine & _
            "Bemerkung: " & existingProject.Remarks, _
            vbYesNo + vbQuestion, "Daten laden")

        If useExisting = vbYes Then
            Set PromptForProjectDetails = existingProject
            Exit Function
        End If
    End If

PromptInputs:
    '--- Prompt for commission number
    Dim commissionInput As Variant
    commissionInput = Application.InputBox( _
        "Kommissionsnummer für " & projectName, _
        "Kommissionsnummer", _
        Type:=2)

    If commissionInput = "False" Or commissionInput = False Then
        '--- User cancelled
        Set PromptForProjectDetails = Nothing
        Exit Function
    End If

    proj.CommissionNumber = CStr(commissionInput)

    '--- Prompt for remarks
    Dim remarksInput As Variant
    remarksInput = Application.InputBox( _
        "Bemerkung für " & projectName, _
        "Bemerkung", _
        Type:=2)

    If remarksInput = "False" Or remarksInput = False Then
        '--- User cancelled
        Set PromptForProjectDetails = Nothing
        Exit Function
    End If

    proj.Remarks = CStr(remarksInput)

    '--- Confirm before saving
    Dim confirmSave As VbMsgBoxResult
    confirmSave = MsgBox( _
        "Soll das Projekt gespeichert werden?" & vbNewLine & vbNewLine & _
        "Projekt: " & proj.projectName & vbNewLine & _
        "Kommission: " & proj.CommissionNumber & vbNewLine & _
        "Bemerkung: " & proj.Remarks, _
        vbYesNo + vbQuestion, "Projekt speichern?")

    If confirmSave = vbNo Then
        '--- Ask again
        GoTo PromptInputs
    End If

    '--- Save to sheet
    If Not SaveProject(proj) Then
        MsgBox "Fehler beim Speichern des Projekts!", vbExclamation
        Set PromptForProjectDetails = Nothing
        Exit Function
    End If

    Set PromptForProjectDetails = proj
End Function

'@Description("Gets the project master worksheet, makes it visible if needed")
'@Returns Worksheet object or Nothing if not found
Public Function GetProjectSheet() As Worksheet
    '@Ignore EmptyStringLiteral
    On Error Resume Next

    Dim ws As Worksheet

    '--- Try to find by name
    Set ws = ThisWorkbook.Worksheets(PROJECT_SHEET_NAME)

    '--- Alternative: Try wsProjekte variable (if defined)
    If ws Is Nothing Then
        Set ws = wsProjekte
    End If

    '--- Make visible if hidden
    If Not ws Is Nothing Then
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
    End If

    Set GetProjectSheet = ws
End Function

'@Description("Loads all projects into a Dictionary for batch processing")
'@Returns Dictionary(ProjectName -> Project)
Public Function LoadAllProjects() As Dictionary
    '@Ignore EmptyStringLiteral
    Dim projectDict As Dictionary
    Set projectDict = New Dictionary

    Dim projectSheet As Worksheet
    Set projectSheet = GetProjectSheet()

    If projectSheet Is Nothing Then
        Set LoadAllProjects = projectDict
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = projectSheet.Cells(projectSheet.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        '--- No data (only header)
        Set LoadAllProjects = projectDict
        Exit Function
    End If

    '--- Read projects
    Dim rowIndex As Long
    Dim proj As Project

    For rowIndex = 2 To lastRow
        Set proj = New Project
        proj.projectName = projectSheet.Cells(rowIndex, 1).value
        proj.CommissionNumber = projectSheet.Cells(rowIndex, 2).value
        proj.Remarks = projectSheet.Cells(rowIndex, 3).value

        If proj.IsValid And Not projectDict.Exists(proj.projectName) Then
            projectDict.Add proj.projectName, proj
        End If
    Next rowIndex

    Set LoadAllProjects = projectDict
End Function


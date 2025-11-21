VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Projekte 
   Caption         =   "Projektauswahl"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "UF_Projekte.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UF_Projekte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("UI.Forms")
'@ModuleDescription("Project selection form - allows users to pick projects from master list")
Option Explicit

Private Const MAIN_PLANNER_FIRST_DAY_COLUMN As Long = 15
Private Const WEEKLY_SHEET_FIRST_DAY_COLUMN As Long = 5

'@Description("UserForm initialization - automatically loads project data when form opens")
'FIX: Added this event handler to ensure projects are loaded when opening via Ribbon
Private Sub UserForm_Initialize()
    Call LoadProjectData
End Sub

'@Description("Refresh button click - reloads project data")
Private Sub CommandButton3_Click()
    Call LoadProjectData
End Sub

'@Description("ListBox double-click - inserts selected project into active cell")
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call WriteProjectToActiveCell
End Sub

'@Description("Writes selected project to the active cell if valid")
Private Sub WriteProjectToActiveCell()
    If ListBox1.ListIndex = -1 Then Exit Sub

    '--- Validate active cell is in a table
    Dim tableName As String
    On Error Resume Next
    tableName = ActiveCell.ListObject.Name
    On Error GoTo 0

    If tableName = vbNullString Then
        Me.Label1.Caption = ActiveCell.Address & " ist ausserhalb des Planers."
        Exit Sub
    End If

    '--- Check if column is a valid day column
    Dim activeColumnNumber As Long
    Dim minimumColumnNumber As Long

    activeColumnNumber = ActiveCell.Column

    '--- Determine minimum column based on sheet type
    Select Case ActiveSheet.Name
        Case "Personalplaner"
            minimumColumnNumber = MAIN_PLANNER_FIRST_DAY_COLUMN
        Case Else
            '--- Assume KW sheet
            minimumColumnNumber = WEEKLY_SHEET_FIRST_DAY_COLUMN
    End Select

    If activeColumnNumber >= minimumColumnNumber Then
        '--- Write project name to cell
        Dim selectedProjectName As String
        selectedProjectName = ListBox1.List(ListBox1.ListIndex, 0)

        ActiveCell.value = selectedProjectName
        Me.Label1.Caption = selectedProjectName & " in Zelle " & ActiveCell.Address & " geschrieben."
    Else
        Me.Label1.Caption = ActiveCell.Address & " ist kein Tag."
    End If
End Sub

'@Description("Loads project data from Projektnummern worksheet into ListBox")
Public Sub LoadProjectData()
    Dim projectSheet As Worksheet
    On Error Resume Next
    Set projectSheet = Worksheets("Projektnummern")
    On Error GoTo 0

    If projectSheet Is Nothing Then
        MsgBox "Blatt 'Projektnummern' nicht gefunden!", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = projectSheet.Cells(projectSheet.Rows.Count, 1).End(xlUp).Row

    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "80;100;120"
        .Clear

        '--- Load data starting from row 2 (row 1 is header)
        Dim rowIndex As Long
        For rowIndex = 2 To lastRow
            .AddItem
            .List(.ListCount - 1, 0) = projectSheet.Cells(rowIndex, 1).value 'Project Name
            .List(.ListCount - 1, 1) = projectSheet.Cells(rowIndex, 2).value 'Commission Number
            .List(.ListCount - 1, 2) = projectSheet.Cells(rowIndex, 3).value 'Remarks
        Next rowIndex
    End With

    Me.Caption = "Projektauswahl"
End Sub

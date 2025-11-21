VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Filter 
   Caption         =   "Filter"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "UF_Filter.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UF_Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("UI.Forms")
'@ModuleDescription("Project filter form - filters visible projects in active sheet")
Option Explicit

Private filterStartColumn As Long
Private cachedUniqueValues As Dictionary

'@Description("Apply filter button - filters rows based on selected values")
Private Sub CommandButton1_Click()
    Call ApplyProjectFilter
    ActiveSheet.Calculate
End Sub

'@Description("Reset filter button - shows all rows")
Private Sub CommandButton2_Click()
    Call ResetAllFilters
End Sub

'@Description("Refresh button - reloads project list")
Private Sub CommandButton3_Click()
    Call LoadFilterData
End Sub

'@Description("ListBox double-click - applies filter immediately")
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ApplyProjectFilter
End Sub

'@Description("Applies filter to hide/show rows based on selected projects")
Private Sub ApplyProjectFilter()
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet

    '--- Collect selected values from ListBox
    Dim selectedValues As Collection
    Set selectedValues = New Collection

    Dim itemIndex As Long
    For itemIndex = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(itemIndex) Then
            selectedValues.Add ListBox1.List(itemIndex)
        End If
    Next itemIndex

    If selectedValues.Count = 0 Then
        MsgBox "Bitte wähle einen oder mehrere Werte aus.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    '--- First, show all rows
    Call ResetFiltersInSheet(targetSheet)

    '--- Apply filter by hiding non-matching rows
    Dim dataTable As ListObject
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim cellValue As String
    Dim rowMatches As Boolean
    Dim selectedValue As Variant

    For Each dataTable In targetSheet.ListObjects
        If Not dataTable.DataBodyRange Is Nothing Then
            For rowIndex = 1 To dataTable.DataBodyRange.Rows.Count
                rowMatches = False

                '--- Check columns from filterStartColumn onwards
                For columnIndex = filterStartColumn To dataTable.ListColumns.Count
                    If columnIndex <= dataTable.DataBodyRange.Columns.Count Then
                        cellValue = Trim$(CStr(dataTable.DataBodyRange.Cells(rowIndex, columnIndex).value))

                        If Len(cellValue) > 0 Then
                            '--- Check if cell value matches any selected value
                            For Each selectedValue In selectedValues
                                If cellValue = selectedValue Then
                                    rowMatches = True
                                    Exit For
                                End If
                            Next selectedValue

                            If rowMatches Then Exit For
                        End If
                    End If
                Next columnIndex

                '--- Hide row if no match found
                If Not rowMatches Then
                    dataTable.DataBodyRange.Rows(rowIndex).EntireRow.Hidden = True
                End If
            Next rowIndex
        End If
    Next dataTable

    Application.ScreenUpdating = True

    Me.Label1.Caption = "Gefiltert nach " & selectedValues.Count & " Wert(en) im aktiven Blatt."
End Sub

'@Description("Resets all filters - shows all rows in all sheets")
Private Sub ResetAllFilters()
    Dim currentSheet As Worksheet
    Dim dataTable As ListObject

    Application.ScreenUpdating = False

    For Each currentSheet In ThisWorkbook.Worksheets
        Call ResetFiltersInSheet(currentSheet)
    Next currentSheet

    Application.ScreenUpdating = True

    Me.Label1.Caption = "Alle Zeilen eingeblendet."
End Sub

'@Description("Resets filters in a specific sheet")
'@Param targetSheet The worksheet to reset filters in
Private Sub ResetFiltersInSheet(ByVal targetSheet As Worksheet)
    Dim dataTable As ListObject

    For Each dataTable In targetSheet.ListObjects
        If Not dataTable.DataBodyRange Is Nothing Then
            dataTable.DataBodyRange.EntireRow.Hidden = False
        End If
    Next dataTable
End Sub

'@Description("Loads unique project values into ListBox based on active sheet")
'@Param startColumn Optional start column (auto-detected if 0)
Public Sub LoadFilterData(Optional ByVal startColumn As Long = 0)
    '--- Auto-detect start column based on sheet type
    If startColumn = 0 Then
        If ActiveSheet.Name Like "KW*" Then
            startColumn = 5
        ElseIf ActiveSheet.Name Like "Personalplaner" Then
            startColumn = 15
        Else
            Exit Sub
        End If
    End If

    filterStartColumn = startColumn

    Me.Caption = "Filter " & ActiveSheet.Name

    '--- Collect unique values using EmployeeService
    Set cachedUniqueValues = EmployeeService.GetUniqueValuesFromListObjects(startColumn)

    '--- Populate ListBox
    ListBox1.Clear

    Dim uniqueKey As Variant
    For Each uniqueKey In cachedUniqueValues.Keys
        ListBox1.AddItem uniqueKey
    Next uniqueKey
End Sub

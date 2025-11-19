Attribute VB_Name = "FilterService"
'@Folder("Services.Filtering")
'@ModuleDescription("Provides filtering functionality for tables using ActiveX ListBox controls")
Option Explicit

'@Description("Applies filter to a table based on ListBox selection")
'@Param targetSheet The worksheet containing the table and ListBox
'@Param listBoxName Name of the ActiveX ListBox control
'@Param columnName Name of the table column to filter
Public Sub ApplyTableFilter( _
        ByVal targetSheet As Worksheet, _
        ByVal listBoxName As String, _
        ByVal columnName As String)

    On Error GoTo ErrorHandler

    Dim dataTable As ListObject
    Set dataTable = targetSheet.ListObjects(1) 'First table in sheet

    Dim listBox As MSForms.ListBox
    Set listBox = targetSheet.OLEObjects(listBoxName).Object

    '--- Collect selected items from ListBox
    Dim selectedItems As Collection
    Set selectedItems = New Collection

    Dim itemIndex As Long
    For itemIndex = 0 To listBox.ListCount - 1
        If listBox.Selected(itemIndex) Then
            '--- Special handling for "Mitarbeiter" column (add wildcard)
            If columnName = "Mitarbeiter" Then
                selectedItems.Add listBox.List(itemIndex) & "*"
            Else
                selectedItems.Add listBox.List(itemIndex)
            End If
        End If
    Next itemIndex

    '--- Reset filter if nothing selected
    If selectedItems.Count = 0 Then
        On Error Resume Next
        dataTable.Range.AutoFilter Field:=dataTable.ListColumns(columnName).Index
        On Error GoTo 0
        Application.Calculate
        Exit Sub
    End If

    '--- Build array with selected values
    Dim filterArray() As Variant
    ReDim filterArray(1 To selectedItems.Count)

    Dim arrayIndex As Long
    For arrayIndex = 1 To selectedItems.Count
        filterArray(arrayIndex) = selectedItems(arrayIndex)
    Next arrayIndex

    '--- Apply filter
    dataTable.Range.AutoFilter _
        Field:=dataTable.ListColumns(columnName).Index, _
        Criteria1:=filterArray, _
        Operator:=xlFilterValues

    Application.Calculate

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Anwenden des Filters:" & vbNewLine & vbNewLine & _
           "Fehler: " & Err.Description, _
           vbExclamation, "Filter-Fehler"
End Sub

'@Description("Resets all filters in a table")
'@Param targetSheet The worksheet containing the table
Public Sub ResetTableFilter(ByVal targetSheet As Worksheet)
    On Error Resume Next

    Dim dataTable As ListObject
    For Each dataTable In targetSheet.ListObjects
        If dataTable.AutoFilter.FilterMode Then
            dataTable.AutoFilter.ShowAllData
        End If
    Next dataTable

    Application.Calculate
End Sub

'@Description("Applies filter when ListBox selection changes (event handler)")
'@Param targetSheet The worksheet containing the controls
'@Param listBoxName Name of the ListBox that triggered the event
'@Param columnName Name of the column to filter
Public Sub OnListBoxChange( _
        ByVal targetSheet As Worksheet, _
        ByVal listBoxName As String, _
        ByVal columnName As String)

    Call ApplyTableFilter(targetSheet, listBoxName, columnName)
End Sub

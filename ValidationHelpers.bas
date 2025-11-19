Attribute VB_Name = "ValidationHelpers"
'@Folder("Services.Utilities")
'@ModuleDescription("Data validation helper functions")
Option Explicit

'@Description("Removes data validation from a range")
'@Param targetRange The range to remove validation from
Public Sub RemoveDataValidation(ByVal targetRange As Range)
    On Error GoTo ErrorHandler

    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean

    '--- Save application state
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents

    '--- Set application state for performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- Validate input
    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 513, "RemoveDataValidation", "Target range is invalid (Nothing)."
    End If

    '--- Remove data validation
    targetRange.Validation.Delete

SafeExit:
    '--- Restore application state
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

ErrorHandler:
    MsgBox "Fehler in 'RemoveDataValidation': " & Err.Number & " - " & Err.Description, _
           vbExclamation, "Fehler beim Entfernen"
    Resume SafeExit
End Sub

'@Description("Checks if a cell has list validation")
'@Param targetCell The cell to check
'@Returns True if cell has list validation
Public Function HasListValidation(ByVal targetCell As Range) As Boolean
    On Error GoTo ErrorHandler
    HasListValidation = (targetCell.Validation.Type = xlValidateList)
    Exit Function

ErrorHandler:
    HasListValidation = False
End Function

'@Description("Applies list validation to a range")
'@Param targetRange The range to apply validation to
'@Param validationList Comma-separated list of valid values
'@Param allowBlank Whether to allow blank values (default True)
Public Sub ApplyListValidation( _
        ByVal targetRange As Range, _
        ByVal validationList As String, _
        Optional ByVal allowBlank As Boolean = True)

    On Error GoTo ErrorHandler

    If targetRange Is Nothing Then Exit Sub
    If Len(Trim$(validationList)) = 0 Then Exit Sub

    '--- Remove existing validation
    Call RemoveDataValidation(targetRange)

    '--- Apply new validation
    With targetRange.Validation
        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:=validationList

        .IgnoreBlank = allowBlank
        .InCellDropdown = True
        .ShowError = False
        .ShowInput = True
        .InputTitle = "Eingabe"
        .InputMessage = "Wähle aus der Liste oder gib frei ein."
    End With

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Anwenden der Validierung: " & Err.Description, _
           vbExclamation, "Validierungsfehler"
End Sub

Attribute VB_Name = "Modul1"
' ============================================================================
' @Description Prüft, ob eine Zelle eine Datenvalidierung vom Typ Liste hat.
' ============================================================================
Private Function HasListValidation(ByVal rng As Range) As Boolean
    On Error GoTo ErrHandler
    HasListValidation = (rng.Validation.Type = xlValidateList)
    Exit Function
ErrHandler:
    HasListValidation = False
End Function

Public Sub ZufallsauswahTeam()
    'On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim lo                             As ListObject
    Dim rngCell                        As Range
    Dim rngDV                          As Range
    Dim i                              As Long
    Dim arrWerte()                     As String
    
    ' Jede Zelle im Tabellenkörper prüfen
    For Each rngCell In Application.Selection
        If HasListValidation(rngCell) Then
            ' Liste der gültigen Werte holen
            
            ' Zufälligen Index wählen
            i = Application.WorksheetFunction.RandBetween(1, 6)
            
            ' Wert einfügen
            rngCell.Value = "Team " & i
        End If
    Next rngCell
    
Cleanup:
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    MsgBox "Fehler bei Zufallsauswahl: " & Err.Description, vbCritical, "Fehler"
    Resume Cleanup
End Sub

Attribute VB_Name = "mDatenüberprüfung"
'@Folder "VALIDATION"
'@ModuleDescription "Werkzeuge zum Entfernen von Datenüberprüfungen."
Option Explicit

'=======================================================================================
'@Description Entfernt die Datenüberprüfung (Data Validation) aus allen Zellen
'             des angegebenen Bereichs.
'
'@Param ZielBereich Range: Bereich, aus dem die Datenüberprüfung entfernt werden soll.
'=======================================================================================
Public Sub EntferneDatenüberprüfung(ByVal ZielBereich As Range)
    On Error GoTo ErrHandler
    
    Dim prevScreenUpdating             As Boolean
    Dim prevEnableEvents               As Boolean
    
    ' Zustände sichern
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    
    ' Zustände setzen
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Eingabevalidierung
    If ZielBereich Is Nothing Then
        Err.Raise vbObjectError + 513, "EntferneDatenüberprüfung", "Der übergebene Bereich ist ungültig (Nothing)."
    End If
    
    ' Datenüberprüfung löschen
    ZielBereich.Validation.Delete
    
SafeExit:
    ' Zustände wiederherstellen
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub
    
ErrHandler:
    ' TODO: Optional Logging ergänzen
    MsgBox "Fehler in 'EntferneDatenüberprüfung': " & Err.Number & " - " & Err.Description, vbExclamation, "Fehler beim Entfernen"
    Resume SafeExit
End Sub


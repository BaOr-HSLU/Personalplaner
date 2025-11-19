Attribute VB_Name = "mFormatierung"
'@Folder "FORMAT"
'@ModuleDescription "Formatiert in allen Zellen eines Bereichs jeweils die erste Zeile (bis zum ersten Zeilenumbruch) fett."
Option Explicit

'=======================================================================================
'@Description Formatiert in allen Zellen eines übergebenen Bereichs die erste Zeile fett.
'             Die übrige Zeichenformatierung der Zellen bleibt unverändert.
'
'@Param ZielBereich Range: Zellen-/Bereichsreferenz, in denen die erste Zeile fett gesetzt werden soll.
'=======================================================================================
Public Sub ErsteZeileImBereichFett(ByVal ZielBereich As Range)
    On Error GoTo ErrHandler
    
    Dim prevScreenUpdating             As Boolean
    Dim prevEnableEvents               As Boolean
    Dim zelle                          As Range
    Dim textInhalt                     As String
    Dim posLF                          As Long
    Dim ersteZeileLaenge               As Long
    
    ' Zustände sichern
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    
    ' Zustände setzen
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Eingabevalidierung
    If ZielBereich Is Nothing Then
        Err.Raise vbObjectError + 513, "ErsteZeileImBereichFett", "Der übergebene Bereich ist ungültig (Nothing)."
    End If
    
    ' Verarbeitung
    For Each zelle In ZielBereich.Cells
        ' Nur auf sichtbare, nicht-leere Zellen mit Text/Anzeige anwenden
        If Not zelle Is Nothing Then
            If Len(zelle.Value2) > 0 Then
                textInhalt = CStr(zelle.Value2)
                
                ' Position des ersten Zeilenumbruchs (ALT+Enter = vbLf)
                posLF = InStr(1, textInhalt, vbNewLine, vbBinaryCompare)
                
                If posLF > 0 Then
                    ersteZeileLaenge = posLF - 1
                Else
                    ersteZeileLaenge = Len(textInhalt)
                End If
                
                If ersteZeileLaenge > 0 Then
                    ' Nur erste Zeile fett setzen, Rest unverändert lassen
                    zelle.Characters().Font.Bold = False
                    zelle.Characters(Start:=1, Length:=ersteZeileLaenge).Font.Bold = True
                End If
            End If
        End If
    Next zelle
    
SafeExit:
    ' Zustände wiederherstellen
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub
    
ErrHandler:
    ' TODO: Bei Bedarf Logging ergänzen
    MsgBox "Fehler in 'ErsteZeileImBereichFett': " & Err.Number & " - " & Err.Description, vbExclamation, "Formatierungsfehler"
    Resume SafeExit
End Sub


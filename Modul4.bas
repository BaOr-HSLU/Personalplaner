Attribute VB_Name = "Modul4"
Sub SendFilteredPDFEmailToAll()
    Dim wsData                         As Worksheet ' Dein gefiltertes Datenblatt
    Dim wsEmails                       As Worksheet ' Tabelle "Personalplaner" mit E-Mail-Adressen
    Dim pdfPath                        As String ' Pfad für den PDF-Export
    Dim cell                           As Range  ' Durchlauf E-Mail-Adressen
    Dim emailList                      As String ' Liste der Empfänger
    Dim lastrow                        As Long   ' Letzte Zeile im Personalplaner
    Dim EMailAddress                   As String
    
    ' Tabellen einstellen
    Set wsData = ActiveSheet                     ' ANPASSEN: Name deines gefilterten Blatts
    Set wsEmails = ThisWorkbook.Sheets("Personalplaner") ' ANPASSEN: Tabelle mit E-Mail-Adressen
    
    ' PDF-Dateipfad festlegen
    pdfPath = Environ("TEMP") & "\" & wsData.Name & ".pdf" ' ANPASSEN: Pfad nach Bedarf
    
    ' Gefiltertes Arbeitsblatt als PDF exportieren
    wsData.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
                               OpenAfterPublish:=False
    
    ' Outlook starten
    Dim outlookApp                     As Object
    Dim mailItem                       As Object
    
    On Error Resume Next                         ' Fehlerbehandlung für den Fall, dass Outlook nicht läuft
    
    ' Erstellt eine neue Instanz der Outlook-Anwendung oder greift auf eine vorhandene zu
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    
    ' Erstellt ein neues E-Mail-Objekt
    Set mailItem = outlookApp.CreateItem(0)      ' 0 steht für olMailItem
    
    ' Letzte Zeile im Personalplaner ermitteln (Annahme: E-Mail ist in Spalte A)
    lastrow = wsEmails.Cells(wsEmails.Rows.Count, "G").End(xlUp).row ' ANPASSEN: Spalte mit E-Mail-Adressen
    
    ' Alle E-Mail-Adressen aus der Tabelle "Personalplaner" sammeln
    emailList = ""
    Dim Key                            As Variant
    Dim dict                           As Dictionary
    Dim lo                             As ListObject
    Set lo = wsData.ListObjects(1)
    Set dict = SammleEindeutigeWerteSchnellRng(lo.ListColumns(2).DataBodyRange, False)
    
    For Each Key In dict.Keys
        EMailAddress = Split(Key, vbNewLine)(2)
        If emailList = "" Then
            emailList = EMailAddress
        Else
            emailList = emailList & ";" & EMailAddress
        End If
    Next Key
    
    ' E-Mail an alle gemeinsam versenden
    With mailItem
        .To = emailList
        .Subject = "Wochenliste " & wsData.Name
        .Body = "Hallo miteinander," & vbCrLf & vbCrLf & _
                "anbei erhaltet ihr die Wochenliste von " & wsData.Name & "." & vbCrLf & vbCrLf
        .Attachments.Add pdfPath
        .Display
        '.Send ' .Display zum Überprüfen, .Send zur direkten Zustellung
    End With
    
    ' Objekte freigeben
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Attribute VB_Name = "EmailService"
'@Folder("Services.Email")
'@ModuleDescription("Email sending service using Outlook automation")
Option Explicit

'@Description("Sends filtered weekly plan as PDF to all visible employees")
Public Sub SendWeeklyPlanPDFToEmployees()
    On Error GoTo ErrorHandler

    Dim activeWorksheet As Worksheet
    Set activeWorksheet = ActiveSheet

    Dim personnelSheet As Worksheet
    Set personnelSheet = ThisWorkbook.Sheets("Personalplaner")

    '--- Export filtered sheet as PDF
    Dim pdfFilePath As String
    pdfFilePath = Environ("TEMP") & "\" & activeWorksheet.Name & ".pdf"

    activeWorksheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfFilePath, _
        OpenAfterPublish:=False

    '--- Collect unique employees from filtered data
    Dim dataTable As ListObject
    Set dataTable = activeWorksheet.ListObjects(1)

    Dim employeeRange As Range
    Set employeeRange = dataTable.ListColumns(2).DataBodyRange

    Dim uniqueEmployees As Dictionary
    Set uniqueEmployees = EmployeeService.GetUniqueValuesFromRange(employeeRange, includeHidden:=False)

    '--- Build email address list
    Dim emailAddressList As String
    emailAddressList = vbNullString

    Dim employeeKey As Variant
    Dim employeeLines() As String
    Dim EmailAddress As String

    For Each employeeKey In uniqueEmployees.Keys
        '--- Parse multi-line employee cell (Name, Phone, Email)
        employeeLines = Split(CStr(employeeKey), vbNewLine)

        If UBound(employeeLines) >= 2 Then
            EmailAddress = Trim$(employeeLines(2))

            If Len(EmailAddress) > 0 And InStr(EmailAddress, "@") > 0 Then
                If emailAddressList <> vbNullString Then
                    emailAddressList = emailAddressList & ";"
                End If
                emailAddressList = emailAddressList & EmailAddress
            End If
        End If
    Next employeeKey

    If emailAddressList = vbNullString Then
        MsgBox "Keine gültigen E-Mail-Adressen gefunden.", vbExclamation, "Keine Empfänger"
        Exit Sub
    End If

    '--- Create Outlook email
    Dim outlookApp As Object
    Dim mailItem As Object

    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler

    If outlookApp Is Nothing Then
        MsgBox "Outlook konnte nicht gestartet werden.", vbCritical, "Outlook-Fehler"
        Exit Sub
    End If

    Set mailItem = outlookApp.CreateItem(0) 'olMailItem

    '--- FIX: Use HTMLBody with UTF-8 charset to properly display German umlauts
    '--- The original .Body property doesn't properly encode UTF-8 characters
    '--- causing umlauts (ä, ö, ü) to display as wrong characters
    Dim emailBodyHTML As String
    emailBodyHTML = "<!DOCTYPE html>" & vbNewLine & _
                    "<html>" & vbNewLine & _
                    "<head>" & vbNewLine & _
                    "<meta charset=""UTF-8"">" & vbNewLine & _
                    "</head>" & vbNewLine & _
                    "<body style=""font-family: Calibri, Arial, sans-serif; font-size: 11pt;"">" & vbNewLine & _
                    "<p>Hallo miteinander,</p>" & vbNewLine & _
                    "<p>anbei erhaltet ihr die Wochenliste von " & activeWorksheet.Name & ".</p>" & vbNewLine & _
                    "<p>Mit freundlichen Grüssen</p>" & vbNewLine & _
                    "</body>" & vbNewLine & _
                    "</html>"

    With mailItem
        .To = emailAddressList
        .Subject = "Wochenliste " & activeWorksheet.Name
        .HTMLBody = emailBodyHTML
        .Attachments.Add pdfFilePath
        .Display 'Show for review (use .Send for automatic sending)
    End With

    Set mailItem = Nothing
    Set outlookApp = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Erstellen der E-Mail:" & vbNewLine & vbNewLine & _
           "Fehler: " & Err.Description, _
           vbCritical, "E-Mail-Fehler"

    Set mailItem = Nothing
    Set outlookApp = Nothing
End Sub

'@Description("Gets Outlook application instance (creates new if needed)")
'@Returns Outlook.Application object or Nothing if failed
Private Function GetOutlookApplication() As Object
    On Error Resume Next

    Set GetOutlookApplication = GetObject(, "Outlook.Application")

    If GetOutlookApplication Is Nothing Then
        Set GetOutlookApplication = CreateObject("Outlook.Application")
    End If
End Function

'@Description("Validates email address format")
'@Param emailAddress The email address to validate
'@Returns True if valid format
Private Function IsValidEmailAddress(ByVal EmailAddress As String) As Boolean
    IsValidEmailAddress = (Len(Trim$(EmailAddress)) > 0 And InStr(EmailAddress, "@") > 0)
End Function

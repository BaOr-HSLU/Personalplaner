VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_CalendarOptions
   Caption         =   "Kalender Optionen"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   OleObjectBlob   =   "UF_CalendarOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_CalendarOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("UI.Forms")
'@ModuleDescription("Calendar creation options dialog")
Option Explicit

Private cancelled As Boolean

'@Description("Gets whether the user cancelled the dialog")
Public Property Get IsCancelled() As Boolean
    IsCancelled = cancelled
End Property

'@Description("Gets the selected start date")
Public Property Get StartDate() As Date
    On Error Resume Next
    StartDate = CDate(txtStartDate.value)
    If Err.Number <> 0 Then StartDate = Date
    On Error GoTo 0
End Property

'@Description("Gets the selected end date")
Public Property Get EndDate() As Date
    On Error Resume Next
    EndDate = CDate(txtEndDate.value)
    If Err.Number <> 0 Then EndDate = Date + 365
    On Error GoTo 0
End Property

'@Description("Gets whether to include holidays and vacations")
Public Property Get IncludeHolidaysVacations() As Boolean
    IncludeHolidaysVacations = chkIncludeHolidaysVacations.value
End Property

'@Description("Gets whether to apply conditional formatting")
Public Property Get ApplyConditionalFormatting() As Boolean
    ApplyConditionalFormatting = chkApplyConditionalFormatting.value
End Property

'@Description("Gets whether to copy formatting from A2 to week cells")
Public Property Get CopyFormattingFromA2() As Boolean
    CopyFormattingFromA2 = chkCopyFormattingFromA2.value
End Property

'@Description("Initializes the form with default values")
Private Sub UserForm_Initialize()
    cancelled = True

    '--- Set default dates
    txtStartDate.value = Format(Date, "dd.mm.yyyy")
    txtEndDate.value = Format(DateAdd("yyyy", 1, Date), "dd.mm.yyyy")

    '--- Set default checkbox values
    chkIncludeHolidaysVacations.value = True
    chkApplyConditionalFormatting.value = True
    chkCopyFormattingFromA2.value = False

    '--- Add labels
    lblStartDate.Caption = "Startdatum (z.B. 01.01.2025):"
    lblEndDate.Caption = "Enddatum (z.B. 31.12.2025):"
    lblOptions.Caption = "Optionen:"

    '--- Add checkbox labels
    chkIncludeHolidaysVacations.Caption = "Feiertage und Schulferien eintragen"
    chkApplyConditionalFormatting.Caption = "Bedingte Formatierung anwenden"
    chkCopyFormattingFromA2.Caption = "Formatierung von A2 auf KW-Zellen kopieren"

    '--- Add button labels
    btnOK.Caption = "OK"
    btnCancel.Caption = "Abbrechen"
End Sub

'@Description("Handles OK button click")
Private Sub btnOK_Click()
    '--- Validate dates
    Dim startDateValue As Date
    Dim endDateValue As Date

    On Error Resume Next
    startDateValue = CDate(txtStartDate.value)
    If Err.Number <> 0 Then
        MsgBox "Ungültiges Startdatum. Bitte Format dd.mm.yyyy verwenden.", vbExclamation
        txtStartDate.SetFocus
        Exit Sub
    End If

    endDateValue = CDate(txtEndDate.value)
    If Err.Number <> 0 Then
        MsgBox "Ungültiges Enddatum. Bitte Format dd.mm.yyyy verwenden.", vbExclamation
        txtEndDate.SetFocus
        Exit Sub
    End If
    On Error GoTo 0

    If endDateValue < startDateValue Then
        MsgBox "Enddatum muss nach dem Startdatum liegen!", vbExclamation
        txtEndDate.SetFocus
        Exit Sub
    End If

    '--- All validations passed
    cancelled = False
    Me.Hide
End Sub

'@Description("Handles Cancel button click")
Private Sub btnCancel_Click()
    cancelled = True
    Me.Hide
End Sub

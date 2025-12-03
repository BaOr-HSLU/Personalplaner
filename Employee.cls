VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Models")
'@ModuleDescription("Represents an employee with contact information")
Option Explicit

Private Type TEmployee
    EmployeeNumber As String
    DisplayName As String
    PhoneNumber As String
    EmailAddress As String
    JobFunction As String
    TeamName As String
    IsSkipped As Boolean
End Type

Private this As TEmployee

'@Description("Gets or sets the employee number")
Public Property Get EmployeeNumber() As String
    EmployeeNumber = this.EmployeeNumber
End Property

Public Property Let EmployeeNumber(ByVal value As String)
    this.EmployeeNumber = value
End Property

'@Description("Gets or sets the display name")
Public Property Get DisplayName() As String
    DisplayName = this.DisplayName
End Property

Public Property Let DisplayName(ByVal value As String)
    this.DisplayName = value
End Property

'@Description("Gets or sets the phone number")
Public Property Get PhoneNumber() As String
    PhoneNumber = this.PhoneNumber
End Property

Public Property Let PhoneNumber(ByVal value As String)
    this.PhoneNumber = value
End Property

'@Description("Gets or sets the email address")
Public Property Get EmailAddress() As String
    EmailAddress = this.EmailAddress
End Property

Public Property Let EmailAddress(ByVal value As String)
    this.EmailAddress = value
End Property

'@Description("Gets or sets the job function")
Public Property Get JobFunction() As String
    JobFunction = this.JobFunction
End Property

Public Property Let JobFunction(ByVal value As String)
    this.JobFunction = value
End Property

'@Description("Gets or sets the team name")
Public Property Get TeamName() As String
    TeamName = this.TeamName
End Property

Public Property Let TeamName(ByVal value As String)
    this.TeamName = value
End Property

'@Description("Gets or sets whether this employee should be skipped in reports")
Public Property Get IsSkipped() As Boolean
    IsSkipped = this.IsSkipped
End Property

Public Property Let IsSkipped(ByVal value As Boolean)
    this.IsSkipped = value
End Property

'@Description("Gets the full name with contact info as multi-line string")
Public Property Get FullContactInfo() As String
    Dim parts As Collection
    Set parts = New Collection

    If Len(this.DisplayName) > 0 Then parts.Add this.DisplayName
    If Len(this.PhoneNumber) > 0 Then parts.Add this.PhoneNumber
    If Len(this.EmailAddress) > 0 Then parts.Add this.EmailAddress

    Dim i As Long
    Dim result As String
    For i = 1 To parts.Count
        If i > 1 Then result = result & vbNewLine
        result = result & parts(i)
    Next i

    FullContactInfo = result
End Property

'@Description("Parses employee data from a multi-line cell value")
'@Param cellValue The cell value with lines: Name, Phone, Email
Public Sub ParseFromCellValue(ByVal cellValue As String)
    If Len(Trim$(cellValue)) = 0 Then Exit Sub

    Dim lines() As String
    lines = Split(cellValue, vbNewLine)

    If UBound(lines) >= 0 Then this.DisplayName = Trim$(lines(0))
    If UBound(lines) >= 1 Then this.PhoneNumber = Trim$(lines(1))
    If UBound(lines) >= 2 Then this.EmailAddress = Trim$(lines(2))
End Sub

'@Description("Validates if employee has a valid email address")
Public Function HasValidEmail() As Boolean
    HasValidEmail = (Len(this.EmailAddress) > 0 And InStr(this.EmailAddress, "@") > 0)
End Function

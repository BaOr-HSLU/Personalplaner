VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Models")
'@ModuleDescription("Represents a project with commission number and remarks")
Option Explicit

Private Type TProject
    projectName As String
    CommissionNumber As String
    Remarks As String
End Type

Private this As TProject

'@Description("Gets or sets the project name")
Public Property Get projectName() As String
    projectName = this.projectName
End Property

Public Property Let projectName(ByVal value As String)
    this.projectName = value
End Property

'@Description("Gets or sets the commission number (Kommissionsnummer)")
Public Property Get CommissionNumber() As String
    CommissionNumber = this.CommissionNumber
End Property

Public Property Let CommissionNumber(ByVal value As String)
    this.CommissionNumber = value
End Property

'@Description("Gets or sets the remarks (Bemerkung)")
Public Property Get Remarks() As String
    Remarks = this.Remarks
End Property

Public Property Let Remarks(ByVal value As String)
    this.Remarks = value
End Property

'@Description("Combines commission number and remarks for storage")
Public Function ToStorageString() As String
    ToStorageString = this.CommissionNumber & ";" & this.Remarks
End Function

'@Description("Parses project data from storage string (CommissionNumber;Remarks)")
Public Sub ParseFromStorageString(ByVal storageString As String)
    Dim parts() As String
    parts = Split(storageString, ";")

    If UBound(parts) >= 0 Then this.CommissionNumber = Trim$(parts(0))
    If UBound(parts) >= 1 Then this.Remarks = Trim$(parts(1))
End Sub

'@Description("Validates if project has minimum required data")
Public Function IsValid() As Boolean
    IsValid = (Len(Trim$(this.projectName)) > 0)
End Function

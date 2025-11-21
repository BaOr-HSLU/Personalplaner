VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AbsenceCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Domain.Models")
'@ModuleDescription("Represents absence codes (F, K, U, WK, etc.) with their descriptions and colors")
'@PredeclaredId
Option Explicit

Private Type TAbsenceCode
    Code As String
    ShortForm As String
    LongForm As String
    ColorRGB As Long
End Type

Private this As TAbsenceCode

'@Description("Creates a new AbsenceCode instance")
Public Function Create(ByVal Code As String, _
                       ByVal ShortForm As String, _
                       ByVal LongForm As String, _
                       ByVal ColorRGB As Long) As AbsenceCode
    With New AbsenceCode
        .Code = Code
        .ShortForm = ShortForm
        .LongForm = LongForm
        .ColorRGB = ColorRGB
        Set Create = .Self
    End With
End Function

Public Property Get Self() As AbsenceCode
    Set Self = Me
End Property

Public Property Get Code() As String
    Code = this.Code
End Property

Public Property Let Code(ByVal value As String)
    this.Code = value
End Property

Public Property Get ShortForm() As String
    ShortForm = this.ShortForm
End Property

Public Property Let ShortForm(ByVal value As String)
    this.ShortForm = value
End Property

Public Property Get LongForm() As String
    LongForm = this.LongForm
End Property

Public Property Let LongForm(ByVal value As String)
    this.LongForm = value
End Property

Public Property Get ColorRGB() As Long
    ColorRGB = this.ColorRGB
End Property

Public Property Let ColorRGB(ByVal value As Long)
    this.ColorRGB = value
End Property

'@Description("Gets all standard absence codes as Dictionary(Code -> AbsenceCode)")
Public Function GetAllCodes() As Dictionary
    Dim codes As Dictionary
    Set codes = New Dictionary

    codes.Add "Fx", AbsenceCode.Create("Fx", "Fx", "Ferien nicht bewilligt", RGB(255, 0, 0))
    codes.Add "F", AbsenceCode.Create("F", "F", "Ferien", RGB(0, 176, 240))
    codes.Add "U", AbsenceCode.Create("U", "U", "Unfall", RGB(255, 192, 0))
    codes.Add "K", AbsenceCode.Create("K", "K", "Krank", RGB(255, 192, 0))
    codes.Add "WK", AbsenceCode.Create("WK", "WK", "Militär", RGB(0, 176, 80))
    codes.Add "S", AbsenceCode.Create("S", "S", "Schule", RGB(224, 237, 201))
    codes.Add "ÜK", AbsenceCode.Create("ÜK", "ÜK", "Überbetr. Kurs", RGB(224, 237, 201))
    codes.Add "T", AbsenceCode.Create("T", "T", "Teilzeit", RGB(191, 191, 191))

    Set GetAllCodes = codes
End Function

'@Description("Gets absence codes as array for validation lists")
Public Function GetCodesArray(Optional ByVal useShortForm As Boolean = True) As Variant
    Dim allCodes As Dictionary
    Set allCodes = GetAllCodes

    Dim result() As String
    ReDim result(0 To allCodes.Count - 1)

    Dim i As Long
    Dim codeKey As Variant
    For Each codeKey In allCodes.Keys
        If useShortForm Then
            result(i) = allCodes(codeKey).ShortForm
        Else
            result(i) = allCodes(codeKey).LongForm
        End If
        i = i + 1
    Next codeKey

    GetCodesArray = result
End Function

'@Description("Checks if a value is a valid absence code")
Public Function IsValidCode(ByVal value As String) As Boolean
    Dim allCodes As Dictionary
    Set allCodes = GetAllCodes
    IsValidCode = allCodes.Exists(Trim$(value))
End Function

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ProjektErstellen 
   Caption         =   "Projekt erfassen"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "UF_ProjektErstellen.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UF_ProjektErstellen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CommandButton1_Click()
    Dim lastRow As Long
    
    If Len(Me.TB_BEZ.value) < 1 Or Len(Me.TB_BEZ.value) < 1 Or Len(Me.TB_BEZ.value) < 1 Then
        If MsgBox("Bist du sicher, dass die Eingaben korrekt sind? Ein oder mehrere Felder sind Leer.", vbYesNo, "Fehlende Eingaben") = vbNo Then Exit Sub
    End If
    
    With wsProjekte
        lastRow = .UsedRange.Rows.Count
        .Cells(lastRow + 1, 1).value = Me.TB_BEZ.value
        .Cells(lastRow + 1, 2).value = Me.TB_KOM.value
        .Cells(lastRow + 1, 3).value = Me.TB_KUN.value
    End With
    
    Me.TB_BEZ.value = ""
    Me.TB_KOM.value = ""
    Me.TB_KUN.value = ""
End Sub

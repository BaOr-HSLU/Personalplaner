VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Filter 
   Caption         =   "Jahresplanung Filtern"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "UF_Filter.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UF_Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder "Personalplaner"
Option Explicit

Private tempCol                        As Long
Public LISTE                           As Dictionary

Private Sub CommandButton1_Click()
    FilterListObjects
    ActiveSheet.Calculate
End Sub

Private Sub CommandButton2_Click()
    Dim ws                             As Worksheet
    Dim lo                             As ListObject
    
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If Not lo.DataBodyRange Is Nothing Then
                lo.DataBodyRange.EntireRow.Hidden = False
            End If
        Next lo
    Next ws
    Application.ScreenUpdating = True
    
    Me.Label1.Caption = "Alle Zeilen eingeblendet."
End Sub

Private Sub CommandButton3_Click()
    LoadData
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    FilterListObjects
End Sub

Private Sub FilterListObjects()
    Dim ws                             As Worksheet
    Dim lo                             As ListObject
    Dim i                              As Long, j As Long, k As Long
    Dim zellenwert                     As String
    Dim gefunden                       As Boolean
    Dim ausgewählteWerte               As Collection
    
    Set ws = ActiveSheet
    
    ' Sammlung der ausgewählten Werte aus ListBox
    Set ausgewählteWerte = New Collection
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            ausgewählteWerte.Add ListBox1.List(i)
        End If
    Next i
    
    If ausgewählteWerte.Count = 0 Then
        MsgBox "Bitte wähle einen oder mehrere Werte aus.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Alle Zeilen im aktiven Blatt einblenden
    For Each lo In ws.ListObjects
        If Not lo.DataBodyRange Is Nothing Then
            lo.DataBodyRange.EntireRow.Hidden = False
        End If
    Next lo
    
    ' Filter anwenden (manuelles Ausblenden)
    For Each lo In ws.ListObjects
        If Not lo.DataBodyRange Is Nothing Then
            For i = 1 To lo.DataBodyRange.Rows.Count
                gefunden = False
                For j = tempCol To lo.ListColumns.Count
                    If j <= lo.DataBodyRange.Columns.Count Then
                        zellenwert = Trim(CStr(lo.DataBodyRange.Cells(i, j).Value))
                        If Len(zellenwert) > 0 Then
                            For k = 1 To ausgewählteWerte.Count
                                If zellenwert = ausgewählteWerte(k) Then
                                    gefunden = True
                                    Exit For
                                End If
                            Next k
                            If gefunden Then Exit For
                        End If
                    End If
                Next j
                If Not gefunden Then
                    lo.DataBodyRange.Rows(i).EntireRow.Hidden = True
                End If
            Next i
        End If
    Next lo
    
    Application.ScreenUpdating = True
    
    Me.Label1.Caption = "Gefiltert nach " & ausgewählteWerte.Count & " Wert(en) im aktiven Blatt."
End Sub

Public Sub LoadData(Optional ByVal startcol As Long = 0)
    If startcol = 0 Then
        If ActiveSheet.Name Like "KW*" Then
            startcol = 5
        ElseIf ActiveSheet.Name Like "Personalplaner" Then
            startcol = 15
        Else
            Exit Sub
        End If
    End If
    
    tempCol = startcol
    
    Me.Caption = "Filter " & ActiveSheet.Name
    
    Set LISTE = SammleEindeutigeWerteSchnell(startcol)
    Dim Key                            As Variant
    With Me.ListBox1
        .Clear
        For Each Key In LISTE.Keys
            .AddItem Key
        Next Key
    End With
End Sub


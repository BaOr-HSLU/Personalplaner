VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Projekte 
   Caption         =   "Jahresplanung Filtern"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "UF_Projekte.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UF_Projekte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'@Folder "Personalplaner"
Option Explicit

Private Sub CommandButton3_Click()
    LoadData
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    WriteProjekt
End Sub

Private Sub WriteProjekt()
    Dim tblName As String
    Dim spaltenNummer As Long
    
    Select Case ActiveSheet.Name
    Case "Personalplaner"
        If ListBox1.listIndex <> -1 Then
            ' Prüfen, ob aktive Zelle in einer Tabelle ist
            On Error Resume Next
            tblName = ActiveCell.ListObject.Name
            On Error GoTo 0
            
            If tblName <> "" Then
                ' Spaltennummer ermitteln (O = 15)
                spaltenNummer = ActiveCell.Column
                
                If spaltenNummer >= 15 Then
                    ' Wert schreiben
                    Me.Label1.Caption = ListBox1.List(ListBox1.listIndex, 0) & " in Zelle " & ActiveCell.Address & " geschrieben."
                    ActiveCell.Value = ListBox1.List(ListBox1.listIndex, 0)
                Else
                    Me.Label1.Caption = ActiveCell.Address & " ist kein Tag."
                End If
            Else
                Me.Label1.Caption = ActiveCell.Address & " ist ausserhalb des Planers."
            End If
        End If
    Case Else
        If ListBox1.listIndex <> -1 Then
            ' Prüfen, ob aktive Zelle in einer Tabelle ist
            On Error Resume Next
            tblName = ActiveCell.ListObject.Name
            On Error GoTo 0
            
            If tblName <> "" Then
                ' Spaltennummer ermitteln (O = 15)
                spaltenNummer = ActiveCell.Column
                
                If spaltenNummer >= 5 Then
                    ' Wert schreiben
                    Me.Label1.Caption = ListBox1.List(ListBox1.listIndex, 0) & " in Zelle " & ActiveCell.Address & " geschrieben."
                    ActiveCell.Value = ListBox1.List(ListBox1.listIndex, 0)
                Else
                    Me.Label1.Caption = ActiveCell.Address & " ist kein Tag."
                End If
            Else
                Me.Label1.Caption = ActiveCell.Address & " ist ausserhalb des Planers."
            End If
        End If

    End Select
End Sub

Public Sub LoadData()
Dim letzteZeile As Long
Dim zeile As Long
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "80;100;120"
        .Clear
        
        ' Letzte Zeile ermitteln
        letzteZeile = Worksheets("Projektnummern").Cells(Rows.Count, 1).End(xlUp).row
        
        ' Daten ab Zeile 2 einlesen (Zeile 1 ist Überschrift)
        For zeile = 2 To letzteZeile
            .AddItem
            .List(.ListCount - 1, 0) = Worksheets("Projektnummern").Cells(zeile, 1).Value
            .List(.ListCount - 1, 1) = Worksheets("Projektnummern").Cells(zeile, 2).Value
            .List(.ListCount - 1, 2) = Worksheets("Projektnummern").Cells(zeile, 3).Value
        Next zeile
    End With
    
    Me.Caption = "Projektauswahl "
End Sub


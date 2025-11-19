Attribute VB_Name = "CustomUI"
'@Folder("CustomUI")
Option Explicit

Private isUILocked           As Boolean
Public myRibbon              As IRibbonUI

#If VBA7 Then
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
#End If

#If VBA7 Then
    Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
    Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If

    Dim objRibbon                As Object
    CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
    Set GetRibbon = objRibbon
    Set objRibbon = Nothing

End Function


'Callback for customUI.onLoad
Sub OnLoad_PERSPLA(ribbon As IRibbonUI)
    #If VBA7 Then
        Dim StoreRibbonPointer As LongPtr
    #Else
        Dim StoreRibbonPointer As Long
    #End If

    'Store Ribbon Object to Public variable
    Set myRibbon = ribbon
    isUILocked = False
    'Store pointer to IRibbonUI in a Named Range within add-in file
    StoreRibbonPointer = ObjPtr(ribbon)
    ThisWorkbook.Names.Add Name:="RibbonID", RefersTo:=StoreRibbonPointer

    Application.StatusBar = " CUSTOM UI | " & "CustomRibbon successfully Loaded"
End Sub

Public Sub RefreshRibbon()
    'PURPOSE: Refresh Ribbon UI

    Dim myRibbon             As Object

    On Error GoTo RestartExcel
    If myRibbon Is Nothing Then
        Set myRibbon = GetRibbon(Replace(ThisWorkbook.Names("RibbonID").RefersTo, "=", vbNullString))
    End If

    'Redo Ribbon Load
    myRibbon.Invalidate
    On Error GoTo 0

    Exit Sub

    'ERROR MESSAGES:
RestartExcel:
    Application.StatusBar = "Please restart Excel for Ribbon UI changes to take affect"
End Sub

'Callback for DASHBOARD getVisible
Sub getVisible_PERSPLA(control As IRibbonControl, ByRef returnedVal)
    Debug.Print control.ID
    Select Case control.ID
    Case "DASHBOARD"
        returnedVal = True
    Case "WOCHENPLAN"
        If ActiveSheet.Name Like "KW*" Then
            returnedVal = True
        Else
            returnedVal = False
        End If
    Case Else
        returnedVal = True
    End Select
End Sub

'Callback for TODAY onAction
Sub onAction_PERSPLA(control As IRibbonControl)
Debug.Print control.ID
Select Case control.ID
Case "TODAY"
    HeuteWählen
Case "ÜBERSICHT"
    HOME
Case "AUSWERTUNG"
    DASHBOARD
Case "DIAGRAMM"
    DIAGRAMM
Case "FILTER"
    UF_Filter.Show 0
Case "PROJEKT"
    UF_Projekte.Show 0
Case "BERECHNEN"
    Application.Calculate
    If ActiveSheet.Name = "Auswertung Mitarbeiter" Then Tabelle8.Auswerten
Case "PROJEKTEINGABE"
    wsProjekte.Visible = xlSheetVisible
    wsProjekte.Activate
    UF_ProjektErstellen.Show 0
Case "WP_SENDEN"
    SendFilteredPDFEmailToAll
Case "WR_ANFORDERUNG"
    WR_Anfordern
Case "WR_ERSTELLEN"
    WR_Erstellen
Case "WR_ERSTELLEN"
    SETTINGS
End Select
End Sub

Private Sub SETTINGS()
    Tabelle5.Visible = xlSheetVisible
    Tabelle5.Activate
End Sub

Private Sub HOME()
    Dim sheet As Worksheet
    Dim i  As Long: i = 1
    
    For i = 1 To ThisWorkbook.Sheets.Count
        With ThisWorkbook.Sheets(i)
        Debug.Print .Name
            Select Case .Name
            Case "Personalplaner"
                .Visible = xlSheetVisible
            Case Else
                .Visible = xlSheetHidden
            End Select
        End With
    Next i
    
    Tabelle3.Activate
    Tabelle8.Visible = xlSheetHidden
    wsProjekte.Visible = xlSheetHidden
    Tabelle5.Visible = xlSheetHidden
End Sub

Private Sub DASHBOARD()
    Tabelle8.Visible = xlSheetVisible
    Tabelle8.Activate
End Sub

Private Sub DIAGRAMM()
    Diagramm1.Visible = xlSheetVisible
    Diagramm1.Activate
End Sub

Private Sub HeuteWählen()
    HOME
    Tabelle3.Cells(10, FindeDatumsspalte(Tabelle3, 10, Now(), 15)).Select
End Sub

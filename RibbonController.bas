Attribute VB_Name = "RibbonController"
'@Folder("UI.Ribbon")
'@ModuleDescription("Manages Custom Ribbon UI interactions and navigation")
Option Explicit

'⚠️ IMPORTANT: Control IDs in this module must match your CustomUI XML ribbon configuration!
'   If you change control IDs here, update the corresponding XML file.

Private ribbonUI As IRibbonUI
Private isRibbonLocked As Boolean

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (pDest As Any, pSource As Any, ByVal byteLength As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (pDest As Any, pSource As Any, ByVal byteLength As Long)
#End If

'@Description("Gets ribbon object from pointer")
#If VBA7 Then
    Private Function GetRibbonFromPointer(ByVal ribbonPointer As LongPtr) As Object
#Else
    Private Function GetRibbonFromPointer(ByVal ribbonPointer As Long) As Object
#End If
    Dim ribbonObject As Object
    CopyMemory ribbonObject, ribbonPointer, LenB(ribbonPointer)
    Set GetRibbonFromPointer = ribbonObject
    Set ribbonObject = Nothing
End Function

'@Description("Callback: Ribbon onLoad - Initialize ribbon reference")
'⚠️ CUSTOMUI XML CALLBACK: onLoad="OnLoad_PersonalPlaner"
Public Sub OnLoad_PersonalPlaner(ByVal ribbon As IRibbonUI)
    #If VBA7 Then
        Dim ribbonPointerStorage As LongPtr
    #Else
        Dim ribbonPointerStorage As Long
    #End If

    Set ribbonUI = ribbon
    isRibbonLocked = False

    '--- Store ribbon pointer in named range
    ribbonPointerStorage = ObjPtr(ribbon)
    ThisWorkbook.Names.Add Name:="RibbonID", RefersTo:=ribbonPointerStorage

    Application.StatusBar = "Custom Ribbon erfolgreich geladen"
End Sub

'@Description("Refreshes the ribbon UI")
Public Sub RefreshRibbon()
    On Error GoTo RestartRequired

    Dim ribbonRef As Object

    If ribbonUI Is Nothing Then
        Set ribbonRef = GetRibbonFromPointer(Replace(ThisWorkbook.Names("RibbonID").RefersTo, "=", vbNullString))
    Else
        Set ribbonRef = ribbonUI
    End If

    ribbonRef.Invalidate
    Exit Sub

RestartRequired:
    Application.StatusBar = "Bitte Excel neu starten für Ribbon-Änderungen"
End Sub

'@Description("Callback: Controls visibility based on active sheet")
'⚠️ CUSTOMUI XML CALLBACK: getVisible="GetControlVisibility"
Public Sub GetControlVisibility(ByVal control As IRibbonControl, ByRef returnVisible As Boolean)
    Select Case control.id
        Case "TabDashboard"
            '⚠️ CUSTOMUI XML: control id="TabDashboard"
            returnVisible = True

        Case "TabWeeklyPlan"
            '⚠️ CUSTOMUI XML: control id="TabWeeklyPlan"
            '--- Only show when in a KW sheet
            returnVisible = (ActiveSheet.Name Like "KW*")

        Case Else
            returnVisible = True
    End Select
End Sub

'@Description("Callback: Button click handler")
'⚠️ CUSTOMUI XML CALLBACK: onAction="OnRibbonButtonClick"
Public Sub OnRibbonButtonClick(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler

    Select Case control.id
        '--- Navigation Buttons ---
        Case "BtnGoToToday"
            '⚠️ CUSTOMUI XML: control id="BtnGoToToday"
            Call NavigateToToday

        Case "BtnShowOverview"
            '⚠️ CUSTOMUI XML: control id="BtnShowOverview"
            Call NavigateToOverview

        Case "BtnShowDashboard"
            '⚠️ CUSTOMUI XML: control id="BtnShowDashboard"
            Call NavigateToDashboard

        Case "BtnShowChart"
            '⚠️ CUSTOMUI XML: control id="BtnShowChart"
            Call NavigateToChart

        '--- Filter & Project Buttons ---
        Case "BtnShowFilter"
            '⚠️ CUSTOMUI XML: control id="BtnShowFilter"
            UF_Filter.Show 0

        Case "BtnShowProjects"
            '⚠️ CUSTOMUI XML: control id="BtnShowProjects"
            UF_Projekte.Show 0

        Case "BtnProjectInput"
            '⚠️ CUSTOMUI XML: control id="BtnProjectInput"
            Call ShowProjectInput

        '--- Calculation Button ---
        Case "BtnRecalculate"
            '⚠️ CUSTOMUI XML: control id="BtnRecalculate"
            Application.Calculate
            If ActiveSheet.Name = "Auswertung Mitarbeiter" Then
                Tabelle8.Auswerten
            End If

        '--- Weekly Report Buttons ---
        Case "BtnSendWeeklyPlan"
            '⚠️ CUSTOMUI XML: control id="BtnSendWeeklyPlan"
            '@Todo Implement SendFilteredPDFEmailToAll
            MsgBox "Funktion 'PDF senden' noch nicht implementiert", vbInformation

        Case "BtnRequestWeeklyReports"
            '⚠️ CUSTOMUI XML: control id="BtnRequestWeeklyReports"
            WeeklyReportService.SendWeeklyReportReminder

        Case "BtnCreateWeeklyReports"
            '⚠️ CUSTOMUI XML: control id="BtnCreateWeeklyReports"
            WeeklyReportService.CreateWeeklyReports

        Case Else
            MsgBox "Unbekannter Button: " & control.id, vbExclamation
    End Select

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Ausführen der Ribbon-Aktion:" & vbNewLine & vbNewLine & _
           "Control: " & control.id & vbNewLine & _
           "Fehler: " & Err.Description, _
           vbCritical, "Ribbon-Fehler"
End Sub

'--- PRIVATE NAVIGATION METHODS ---

'@Description("Navigates to today's date column in main planner")
Private Sub NavigateToToday()
    Call NavigateToOverview

    Dim todayColumn As Long
    todayColumn = DateHelpers.FindDateColumn(Tabelle3, 10, Now(), 15)

    If todayColumn > 0 Then
        Tabelle3.Cells(10, todayColumn).Select
    Else
        MsgBox "Heutiges Datum nicht im Kalender gefunden", vbInformation
    End If
End Sub

'@Description("Shows overview (main Personalplaner)")
Private Sub NavigateToOverview()
    Dim currentSheet As Worksheet

    '--- Hide all sheets except Personalplaner
    For Each currentSheet In ThisWorkbook.Worksheets
        Select Case currentSheet.Name
            Case "Personalplaner"
                currentSheet.Visible = xlSheetVisible
            Case Else
                currentSheet.Visible = xlSheetHidden
        End Select
    Next currentSheet

    '--- Ensure specific sheets stay hidden
    Tabelle3.Activate
    Tabelle8.Visible = xlSheetHidden
    wsProjekte.Visible = xlSheetHidden
    Tabelle5.Visible = xlSheetHidden
End Sub

'@Description("Shows dashboard (Auswertung Mitarbeiter)")
Private Sub NavigateToDashboard()
    Tabelle8.Visible = xlSheetVisible
    Tabelle8.Activate
End Sub

'@Description("Shows chart sheet")
Private Sub NavigateToChart()
    Diagramm1.Visible = xlSheetVisible
    Diagramm1.Activate
End Sub

'@Description("Shows project input form")
Private Sub ShowProjectInput()
    wsProjekte.Visible = xlSheetVisible
    wsProjekte.Activate
    UF_ProjektErstellen.Show 0
End Sub

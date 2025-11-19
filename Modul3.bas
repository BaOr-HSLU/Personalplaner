Attribute VB_Name = "Modul3"
'@Folder("TOOLS")
'@ModuleDescription "Zählt Stunden für einen Mitarbeiter mit mehreren möglichen Kriterien."
Option Explicit

Public Function Stundenzähler(ByVal Mitarbeiter As String, ByVal strFilter As String) As Double
    Application.Volatile True
    Application.ScreenUpdating = False
    Application.StatusBar = "Zählt " & strFilter & "Tage für " & Mitarbeiter
    
    On Error GoTo ErrHandler
    
    Dim startcol                       As Long
    Dim endcol                         As Long
    Dim ws                             As Worksheet
    Dim Target                         As Range
    Dim mitarbeiterRow                 As Long
    Dim kriterium                      As Variant
    Dim summe                          As Double
    Dim Filter()                       As String
    
    Set Target = Application.Caller
    Set ws = Target.Parent
    
    ' Spalten anhand Datum suchen
    startcol = FindeDatumsspalte(Tabelle3, 10, ws.Range("E4").Value)
    endcol = FindeDatumsspalte(Tabelle3, 10, ws.Range("F4").Value)
    
    ' Mitarbeiterzeile suchen
    Dim rngFound                       As Range
    Set rngFound = Tabelle3.Range("G:G").Find(What:=Mitarbeiter, LookAt:=xlWhole)
    If rngFound Is Nothing Then
        Stundenzähler = 0
        Exit Function
    End If
    mitarbeiterRow = rngFound.row
    
    ' Bereich definieren
    Dim rngCheck                       As Range
    Set rngCheck = Tabelle3.Range(Tabelle3.Cells(mitarbeiterRow, startcol), Tabelle3.Cells(mitarbeiterRow, endcol))
    
    ' Alle Filterkriterien durchlaufen und summieren
    summe = 0
    
    Select Case strFilter
        Case "Frei"
            summe = WorksheetFunction.CountBlank(rngCheck)
            
        Case "Projekt"
            Filter = Split("F,Fx,S,ÜK,U,K,WK,T", ",")
            summe = 0
            For Each kriterium In Filter
                summe = summe + WorksheetFunction.CountIf(rngCheck, kriterium)
            Next kriterium
            summe = WorksheetFunction.CountA(rngCheck) - summe
        Case Else
            Filter = Split(strFilter, ";")
            summe = 0
            For Each kriterium In Filter
                summe = summe + WorksheetFunction.CountIf(rngCheck, kriterium)
            Next kriterium
    End Select
    
    Stundenzähler = summe
CleanExit:
    Exit Function
    
    Application.ScreenUpdating = True
    
ErrHandler:
    Stundenzähler = CVErr(xlErrValue)
    Resume CleanExit
End Function


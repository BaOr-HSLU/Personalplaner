Attribute VB_Name = "mAuslastung"
'@Folder "FORMELN"
'@ModuleDescription "UDFs für robustes Finden von Datumsspalten (Header) und Auslesen des letzten Werts mit optionalem Spalten-Offset."
Option Explicit

'==================================================================================================
'@Description Liefert den Wert aus der letzten Datenzeile in der Spalte, deren Header einem Datum
'             entspricht; die Zielspalte kann optional per Offset relativ verschoben werden.
'
'@Param datum         Date  : Datum, das in der Kopfzeile gesucht wird (exakte Übereinstimmung, Zeitanteil wird ignoriert).
'@Param offset        Long  : Spaltenversatz relativ zur gefundenen Datumsspalte (0 = genau die Datumsspalte).
'@Param HeaderRow     Long  : Zeilennummer der Kopfzeile (Standard 10).
'@Param DataStartRow  Long  : Erste Datenzeile unterhalb des Headers (Standard 11).
'@Param DataAnchorCol Long  : Spalte, in der die letzte Datenzeile ermittelt wird (Standard 1 = Spalte A).
'
'@Return Variant: Zellenwert oder Fehler (#NV, #WERT!, #BEZUG!).
'==================================================================================================
Public Function VerweisMABAuslastungTotal( _
       ByVal Datum As Date, _
       Optional ByVal offset As Long = 0, _
       Optional ByVal HeaderRow As Long = 10, _
       Optional ByVal DataStartRow As Long = 15, _
       Optional ByVal DataAnchorCol As Long = 1) As Double
    
    On Error GoTo ErrHandler
    
    Dim ws                             As Worksheet
    Dim lastrow                        As Long
    Dim dateCol                        As Long
    Dim usedLastCol                    As Long
    
    '--- Zielblatt per CodeName (anpassen, falls nötig)
    Set ws = Tabelle3
    
    '--- Eingabeprüfung
    If HeaderRow < 1 Or DataStartRow < 1 Or DataStartRow <= HeaderRow Then
        VerweisMABAuslastungTotal = CVErr(xlErrValue)
        Exit Function
    End If
    If DataAnchorCol < 1 Or DataAnchorCol > ws.Columns.Count Then
        VerweisMABAuslastungTotal = CVErr(xlErrRef)
        Exit Function
    End If
    
    '--- Letzte belegte Datenzeile in Anker-Spalte
    lastrow = ws.Cells(ws.Rows.Count, DataAnchorCol).End(xlUp).row
    If lastrow < DataStartRow Then
        VerweisMABAuslastungTotal = CVErr(xlErrNA) ' Keine Datenzeilen vorhanden
        Exit Function
    End If
    
    '--- Begrenzung der Header-Suche auf genutzte Spalten
    usedLastCol = GetUsedLastColumn(ws)
    If usedLastCol = 0 Then
        VerweisMABAuslastungTotal = CVErr(xlErrNA)
        Exit Function
    End If
    
    '--- Datumsspalte robust ermitteln (ignoriert Zeitanteil; kann mit Text-Headern umgehen)
    dateCol = FindeDatumsspalte(ws, HeaderRow, Datum, 1, usedLastCol)
    If dateCol = 0 Then
        VerweisMABAuslastungTotal = CVErr(xlErrNA) ' Datum im Header nicht gefunden
        Exit Function
    End If
    
    '--- Offset anwenden und Grenzen prüfen
    dateCol = dateCol + offset
    If dateCol < 1 Or dateCol > ws.Columns.Count Then
        VerweisMABAuslastungTotal = CVErr(xlErrRef) ' Offset führt aus dem Blatt
        Exit Function
    End If
    
    '--- Ergebnis (letzte Datenzeile in der Zielspalte)
    VerweisMABAuslastungTotal = ws.Cells(lastrow + 1, dateCol).Value
    'Debug.Print ws.Cells(lastRow + 1, dateCol).Value, lastRow, dateCol, ws.Cells(lastRow + 1, dateCol).Address
    Exit Function
    
ErrHandler:
    VerweisMABAuslastungTotal = CVErr(xlErrValue)
End Function

'==================================================================================================
'@Description Ermittelt die letzte verwendete Spalte des Blatts (0, falls das Blatt leer ist).
'
'@Param ws Worksheet: Arbeitsblatt.
'@Return Long: Letzte verwendete Spalte oder 0.
'==================================================================================================
Private Function GetUsedLastColumn(ByVal ws As Worksheet) As Long
    On Error GoTo Fail
    Dim lastCol                        As Variant
    GetUsedLastColumn = ws.Columns.Count
    Exit Function
Fail:
    GetUsedLastColumn = 0
End Function

Public Function AbwesendeMAB(Datum As Date) As Long
    Dim col                            As Long: col = FindeDatumsspalte(Tabelle3, 10, Datum, 15)
    AbwesendeMAB = ZaehleCodes(Intersect(Tabelle3.Columns(col), Tabelle3.UsedRange))
End Function

'==================================================================================================
'@Description Zählt, wie viele Zellen in einem Bereich einen der angegebenen Werte enthalten.
'             Beispiel: =ZaehleCodes(A1:A100) zählt alle Zellen mit F, U, K, WK, S oder ÜK.
'
'@Param ZielBereich Range: Der Zellbereich, der geprüft werden soll.
'@Return Long: Anzahl der Zellen, die einen der gesuchten Werte enthalten.
'==================================================================================================
Public Function ZaehleCodes(ByVal ZielBereich As Range) As Long
    'On Error GoTo ErrHandler
    
    Dim zelle                          As Range
    Dim zaehler                        As Long
    Dim arrCodes                       As Variant
    Dim i                              As Long
    
    ' Gesuchte Werte definieren
    arrCodes = Array("F", "U", "K", "WK", "S", "ÜK", "T")
    
    zaehler = 0
    For Each zelle In ZielBereich.Cells
        If Not IsEmpty(zelle.Value) Then
            For i = LBound(arrCodes) To UBound(arrCodes)
                ' Exakter Vergleich, Groß/Kleinschreibung ignoriert
                If StrComp(Trim$(CStr(zelle.Value)), arrCodes(i), vbTextCompare) = 0 Then
                    zaehler = zaehler + 1
                    Exit For                     ' nicht mehrfach zählen
                End If
            Next i
        End If
    Next zelle
    
    ZaehleCodes = zaehler
    Exit Function
    
ErrHandler:
    ZaehleCodes = CVErr(xlErrValue)
End Function



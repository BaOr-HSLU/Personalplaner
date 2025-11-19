Attribute VB_Name = "mBerechnung"
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
'@Description Findet die Spaltennummer eines Datums in der Headerzeile. Arbeitet robust gegen
'             Formatunterschiede: echte Datumswerte, Datum mit Zeitanteil, oder als Text (z. B. "01.02.2025").
'
'@Param ws           Worksheet : Arbeitsblatt mit Header.
'@Param HeaderRow    Long      : Zeile des Headers.
'@Param Suchdatum    Date      : Gesuchtes Datum (Zeitanteil wird ignoriert).
'@Param ErsteSpalte  Long      : Erste zu prüfende Spalte (Standard 1).
'@Param LetzteSpalte Long      : Letzte zu prüfende Spalte (0 = automatisch ermitteln).
'
'@Return Long: Spaltennummer (1-basiert) oder 0 wenn nicht gefunden.
'==================================================================================================
Public Function FindeDatumsspalte( _
       ByVal ws As Worksheet, _
       ByVal HeaderRow As Long, _
       ByVal Suchdatum As Date, _
       Optional ByVal ErsteSpalte As Long = 1, _
       Optional ByVal LetzteSpalte As Long = 0) As Long
    
    On Error GoTo CleanFail
    
    Dim hdrRange                       As Range
    Dim m                              As Variant
    Dim suchSerial                     As Double
    Dim c                              As Range
    Dim autoLastCol                    As Long
    Dim txtTarget1                     As String
    Dim txtTarget2                     As String
    
    '--- Grenzen bestimmen
    If LetzteSpalte = 0 Then
        autoLastCol = GetUsedLastColumn(ws)
        If autoLastCol = 0 Then GoTo CleanFail
        LetzteSpalte = autoLastCol
    End If
    If ErsteSpalte < 1 Then ErsteSpalte = 1
    If LetzteSpalte > ws.Columns.Count Then LetzteSpalte = ws.Columns.Count
    If LetzteSpalte < ErsteSpalte Then GoTo CleanFail
    
    '--- Headerbereich
    Set hdrRange = ws.Range(ws.Cells(HeaderRow, ErsteSpalte), ws.Cells(HeaderRow, LetzteSpalte))
    
    '--- Zielwerte vorbereiten (Zeitanteil ignorieren)
    suchSerial = Int(CDbl(Suchdatum))
    txtTarget1 = Format$(Suchdatum, "dd.mm.yyyy")
    txtTarget2 = Format$(Suchdatum, "d.m.yyyy")
    
    '--- 1) Direkter MATCH auf numerische Datumssereien (funktioniert, wenn Header echte Datumswerte sind)
    m = Application.Match(suchSerial, hdrRange, 0)
    If Not IsError(m) Then
        FindeDatumsspalte = ErsteSpalte + CLng(m) - 1
        Exit Function
    End If
    
    '--- 2) MATCH auf Textrepräsentation (falls Header z. B. "01.02.2025" als Text enthält)
    m = Application.Match(txtTarget1, hdrRange, 0)
    If Not IsError(m) Then
        FindeDatumsspalte = ErsteSpalte + CLng(m) - 1
        Exit Function
    End If
    m = Application.Match(txtTarget2, hdrRange, 0)
    If Not IsError(m) Then
        FindeDatumsspalte = ErsteSpalte + CLng(m) - 1
        Exit Function
    End If
    
    '--- 3) Manuelle Schleife: berücksichtigt Zeitanteile und Text, der in ein Datum konvertierbar ist
    For Each c In hdrRange.Cells
        If Not IsError(c.Value2) Then
            If IsDate(c.Value) Then
                If Int(CDbl(CDate(c.Value))) = suchSerial Then
                    FindeDatumsspalte = c.Column
                    Exit Function
                End If
            Else
                ' Text -> versuchen zu parsen
                If LenB(c.Value2) > 0 Then
                    If IsDate(CStr(c.Value2)) Then
                        If Int(CDbl(CDate(CStr(c.Value2)))) = suchSerial Then
                            FindeDatumsspalte = c.Column
                            Exit Function
                        End If
                    Else
                        ' Direkter Textvergleich auf gängige Darstellungen
                        If CStr(c.Value2) = txtTarget1 Or CStr(c.Value2) = txtTarget2 Then
                            FindeDatumsspalte = c.Column
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next c
    
    '--- Nicht gefunden
    FindeDatumsspalte = 0
    Exit Function
    
CleanFail:
    FindeDatumsspalte = 0
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

' ===========================================================================
' Öffentliche Funktionen
' ===========================================================================

'@Description "Berechnet die Auslastung: Anteil nicht-ausgeschlossener Mitarbeiter in der Spalte, bezogen auf alle verfügbaren Mitarbeiter."
'@Param rngAusschluss "Spalten-/Listenbereich mit auszuschließenden Werten (z. B. 'Ferien', 'Militär', ...)."
'@Return "Double in [0..1]; 0 wenn keine Mitarbeiter verfügbar sind."
Public Function AuslastungMitAusschluss( _
       ByVal rngAusschluss As Range, _
       Optional ByVal abteilung = False) As Double
    
    Application.Volatile True
    
    On Error GoTo ErrHandler
    
    Dim ws                             As Worksheet
    Set ws = Application.Caller.Worksheet
    
    Dim lo                             As ListObject
    Set lo = ws.ListObjects(1)                   ' erste Tabelle im Blatt
    
    ' Tabellenüberschrift in derselben Spalte wie Formel bestimmen
    Dim colIndex                       As Long
    colIndex = Application.Caller.Column - lo.Range.Columns(1).Column + 1
    If colIndex < 1 Or colIndex > lo.ListColumns.Count Then GoTo SafeExit
    
    Dim colName                        As String
    colName = lo.HeaderRowRange.Cells(1, colIndex).Value
    
    Dim rngTag                         As Range
    Set rngTag = lo.ListColumns(colName).DataBodyRange
    
    Dim rngMitarbeiter                 As Range
    Set rngMitarbeiter = lo.ListColumns("Mitarbeiter").DataBodyRange
    
    ' Berechnung durchführen
    If abteilung Then
        AuslastungMitAusschluss = BerechneAuslastungAlle(rngTag, rngMitarbeiter, rngAusschluss)
    Else
        AuslastungMitAusschluss = BerechneAuslastung(rngTag, rngMitarbeiter, rngAusschluss)
    End If
    Exit Function
    
SafeExit:
    AuslastungMitAusschluss = 0#
    Exit Function
    
ErrHandler:
    Resume SafeExit
End Function

'@Description "Zählt die Anzahl verfügbarer Mitarbeiter in der Spalte (nicht in Ausschlussliste)."
'@Param rngAusschluss "Spalten-/Listenbereich mit auszuschließenden Werten."
'@Return "Long; Anzahl Mitarbeiter, die verfügbar sind."
Public Function VerfuegbareMitarbeiter( _
       ByVal rngAusschluss As Range, _
       Optional ByVal abteilung = False) As Long
    
    Application.Volatile True
    
    On Error GoTo ErrHandler
    
    Dim ws                             As Worksheet
    Set ws = Application.Caller.Worksheet
    
    Dim lo                             As ListObject
    Set lo = ws.ListObjects(1)
    
    ' Spaltenüberschrift anhand der Formelposition bestimmen
    Dim colIndex                       As Long
    colIndex = Application.Caller.Column - lo.Range.Columns(1).Column + 1
    If colIndex < 1 Or colIndex > lo.ListColumns.Count Then GoTo SafeExit
    
    Dim colName                        As String
    colName = lo.HeaderRowRange.Cells(1, colIndex).Value
    
    Dim rngTag                         As Range
    Set rngTag = lo.ListColumns(colName).DataBodyRange
    
    Dim rngMitarbeiter                 As Range
    Set rngMitarbeiter = lo.ListColumns("Mitarbeiter").DataBodyRange
    
    If abteilung Then
        VerfuegbareMitarbeiter = BerechneAlle(rngTag, rngMitarbeiter, rngAusschluss)
    Else
        VerfuegbareMitarbeiter = BerechneVerfuegbare(rngTag, rngMitarbeiter, rngAusschluss)
    End If
    Exit Function
    
SafeExit:
    VerfuegbareMitarbeiter = 0
    Exit Function
    
ErrHandler:
    Resume SafeExit
End Function

' ===========================================================================
' Interne Hilfsfunktionen
' ===========================================================================

Private Function BerechneAuslastung( _
        ByVal rngTag As Range, _
        ByVal rngMitarbeiter As Range, _
        ByVal rngAusschluss As Range) As Double
    
    Dim dictExcl                       As Object
    Set dictExcl = CreateObject("Scripting.Dictionary")
    dictExcl.CompareMode = vbTextCompare
    
    Dim c                              As Range, keyVal As String
    For Each c In rngAusschluss.Cells
        If Not IsError(c.Value2) Then
            keyVal = Trim$(CStr(c.Value2))
            If Len(keyVal) > 0 Then
                If Not dictExcl.Exists(keyVal) Then dictExcl.Add keyVal, True
            End If
        End If
    Next c
    
    Dim i                              As Long
    Dim zelleTag                       As Range, zelleMit As Range
    Dim zaehler                        As Long, nenner As Long
    
    For i = 1 To rngTag.Rows.Count
        Set zelleTag = rngTag.Cells(i, 1)
        Set zelleMit = rngMitarbeiter.Cells(i, 1)
        
        If Not zelleTag.EntireRow.Hidden Then
            If IsCellNonEmpty(zelleMit) Then
                keyVal = Trim$(SafeString(zelleTag.Value2))
                ' Nur zählen, wenn Mitarbeiter nicht in Ausschlussliste
                If Len(keyVal) > 0 And Not dictExcl.Exists(keyVal) Then
                    zaehler = zaehler + 1
                    nenner = nenner + 1
                ElseIf Len(keyVal) = 0 Then
                    nenner = nenner + 1
                End If
            End If
        End If
    Next i
    
    If nenner = 0 Then
        BerechneAuslastung = 0#
    Else
        BerechneAuslastung = zaehler / nenner
    End If
End Function

Private Function BerechneAuslastungAlle( _
        ByVal rngTag As Range, _
        ByVal rngMitarbeiter As Range, _
        ByVal rngAusschluss As Range) As Double
    
    Dim dictExcl                       As Object
    Set dictExcl = CreateObject("Scripting.Dictionary")
    dictExcl.CompareMode = vbTextCompare
    
    Dim c                              As Range, keyVal As String
    For Each c In rngAusschluss.Cells
        If Not IsError(c.Value2) Then
            keyVal = Trim$(CStr(c.Value2))
            If Len(keyVal) > 0 Then
                If Not dictExcl.Exists(keyVal) Then dictExcl.Add keyVal, True
            End If
        End If
    Next c
    
    Dim i                              As Long
    Dim zelleTag                       As Range, zelleMit As Range
    Dim zaehler                        As Long, nenner As Long
    
    For i = 1 To rngTag.Rows.Count
        Set zelleTag = rngTag.Cells(i, 1)
        Set zelleMit = rngMitarbeiter.Cells(i, 1)
        
        If IsCellNonEmpty(zelleMit) Then
            keyVal = Trim$(SafeString(zelleTag.Value2))
            ' nur zählen, wenn Mitarbeiter nicht in Ausschlussliste
            If Len(keyVal) > 0 And Not dictExcl.Exists(keyVal) Then
                zaehler = zaehler + 1
                nenner = nenner + 1
            ElseIf Len(keyVal) = 0 Then
                nenner = nenner + 1
            End If
        End If
    Next i
    
    If nenner = 0 Then
        BerechneAuslastungAlle = 0#
    Else
        BerechneAuslastungAlle = zaehler / nenner
    End If
End Function

Private Function BerechneVerfuegbare( _
        ByVal rngTag As Range, _
        ByVal rngMitarbeiter As Range, _
        ByVal rngAusschluss As Range) As Long
    
    Dim dictExcl                       As Object
    Set dictExcl = CreateObject("Scripting.Dictionary")
    dictExcl.CompareMode = vbTextCompare
    
    Dim c                              As Range, keyVal As String
    For Each c In rngAusschluss.Cells
        If Not IsError(c.Value2) Then
            keyVal = Trim$(CStr(c.Value2))
            If Len(keyVal) > 0 Then
                If Not dictExcl.Exists(keyVal) Then dictExcl.Add keyVal, True
            End If
        End If
    Next c
    
    Dim i                              As Long
    Dim zelleTag                       As Range, zelleMit As Range
    Dim verfuegbar                     As Long
    
    For i = 1 To rngTag.Rows.Count
        Set zelleTag = rngTag.Cells(i, 1)
        Set zelleMit = rngMitarbeiter.Cells(i, 1)
        
        If Not zelleTag.EntireRow.Hidden Then
            If IsCellNonEmpty(zelleMit) Then
                keyVal = Trim$(SafeString(zelleTag.Value2))
                If Len(keyVal) = 0 Then          'Or Not dictExcl.Exists(keyVal) Then
                    verfuegbar = verfuegbar + 1
                End If
            End If
        End If
    Next i
    
    BerechneVerfuegbare = verfuegbar
End Function

Private Function BerechneAlle( _
        ByVal rngTag As Range, _
        ByVal rngMitarbeiter As Range, _
        ByVal rngAusschluss As Range) As Long
    
    Dim dictExcl                       As Object
    Set dictExcl = CreateObject("Scripting.Dictionary")
    dictExcl.CompareMode = vbTextCompare
    
    Dim c                              As Range, keyVal As String
    For Each c In rngAusschluss.Cells
        If Not IsError(c.Value2) Then
            keyVal = Trim$(CStr(c.Value2))
            If Len(keyVal) > 0 Then
                If Not dictExcl.Exists(keyVal) Then dictExcl.Add keyVal, True
            End If
        End If
    Next c
    
    Dim i                              As Long
    Dim zelleTag                       As Range, zelleMit As Range
    Dim verfuegbar                     As Long
    
    For i = 1 To rngTag.Rows.Count
        Set zelleTag = rngTag.Cells(i, 1)
        Set zelleMit = rngMitarbeiter.Cells(i, 1)
        
        If IsCellNonEmpty(zelleMit) Then
            keyVal = Trim$(SafeString(zelleTag.Value2))
            If Len(keyVal) = 0 Or Not dictExcl.Exists(keyVal) Then
                verfuegbar = verfuegbar + 1
            End If
        End If
    Next i
    
    BerechneAlle = verfuegbar
End Function

Private Function IsCellNonEmpty(ByVal Target As Range) As Boolean
    On Error Resume Next
    IsCellNonEmpty = (Len(Trim$(SafeString(Target.Value2))) > 0)
End Function

Private Function SafeString(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        SafeString = vbNullString
    Else
        SafeString = CStr(v)
    End If
End Function



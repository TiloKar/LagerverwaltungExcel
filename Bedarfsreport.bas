Attribute VB_Name = "Bedarfsreport"
Public Sub ReportBedarf()
    
    'to do Call holeDatenbank, dann Liste im Arbeitsspeicher erstellen und danach in report plotten
    
    Dim i As Integer
    Dim k As Integer
    Dim reportzeile As Integer
    reportzeile = 2
    'Dim letzteprojektzeile As Integer
    Dim ersterTreffer As Boolean
    Dim ausgabezeile As Integer
    
    ersterTreffer = True
    fund = False
    On Error GoTo FehlerOpen
    Workbooks.Open Filename:=a & b, ReadOnly:=True, Password:=pwlager
    On Error GoTo FehlerNachOpen
    'Sind Projekte vorhanden
    If ActiveWorkbook.Sheets.Count > 1 Then
        Dim Projektblatt As Worksheet
        Dim NeueMappe As Worksheet
        Dim treffer As Object
        Dim offsetSpalte As Integer
        Dim anzahlProjekte As Integer
        'wieviele Projekte gibt es
        For i = 2 To Workbooks(b).Sheets.Count
            Set Projektblatt = Workbooks(b).Worksheets(i)
            If Projektblatt.UsedRange.Rows.Count > 1 Then
                For k = 2 To Projektblatt.UsedRange.Rows.Count
                    If Not IsNumeric(Projektblatt.Cells(k, 7).Value) Then
                        MsgBox "Projekt: " & Projektblatt.Name & " Bestand in Zeile " & k & " kann nicht als Zahl interpretiert werden"
                        Exit For
                    ElseIf Not IsNumeric(Projektblatt.Cells(k, 8).Value) Then
                        MsgBox "Projekt: " & Projektblatt.Name & " Bedarf in Zeile " & k & " kann nicht als Zahl interpretiert werden"
                        Exit For
                    ElseIf CSng(Projektblatt.Cells(k, 8).Value) > CSng(Projektblatt.Cells(k, 7).Value) Then
                        If ersterTreffer Then       ' kopfzeile anlegen
                            ersterTreffer = False
                            Workbooks.Add (a + "\Report Bedarf.xltm")
                            Set NeueMappe = ActiveWorkbook.Sheets(1)
                            NeueMappe.Cells(1, 9).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss") 'aktuelle uhrzeit im kopf
                            anzahlProjekte = 0
                        End If
                        'gibts schon oder nicht?
                            
                 
                        Set Spalte = Nothing    'spalte mit projekt finden
                        Set Spalte = NeueMappe.Rows(2).Find(what:=Projektblatt.Name, lookat:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=False)
                        If Spalte Is Nothing Then
                            offsetSpalte = 11 + 5 * anzahlProjekte
                            anzahlProjekte = anzahlProjekte + 1
                          
                        Else
                            offsetSpalte = Spalte.Column
                        End If
                               
                        Set treffer = Nothing
                        Set treffer = NeueMappe.Columns(1).Find(what:=Projektblatt.Cells(k, 1).Value, lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
                        
                        If treffer Is Nothing Then 'ean noch nicht angelegt
                            NeueMappe.Rows(3).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
                            Projektblatt.Range(Projektblatt.Cells(k, 1), Projektblatt.Cells(k, 6)).Copy Destination:=NeueMappe.Cells(3, 1) 'Artikeldaten aus projektmappe erstmals in report kopieren
                            NeueMappe.Cells(3, 7).Value = 0 'Summe auf 0 initialisieren
                            NeueMappe.Cells(3, 10).Value = Projektblatt.Cells(k, 10).Value 'letzte Bedarfsmeldung für diesen artikel übernehmen
                            
                            ausgabezeile = 3
                            Set treffer = Nothing
                            Set treffer = Workbooks(b).Worksheets(1).Columns(1).Find(what:=Projektblatt.Cells(k, 1).Value, lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
                            If Not (treffer Is Nothing) Then 'in stammdaten gefunden
                                NeueMappe.Cells(3, 8).Value = Workbooks(b).Worksheets(1).Cells(treffer.Row, 9).Value 'Lageranzahl aus sstammdaten setzen setzen
                                NeueMappe.Cells(3, 9).Value = Workbooks(b).Worksheets(1).Cells(treffer.Row, 15).Value 'Bestellt einmalig auf den in der stammdatenliste hinterlegten wert setzen
                            End If
                            
                           
                        Else 'ean schon da, addieren und projekt hinweis
                            ausgabezeile = treffer.Row
                        End If
                        
                        plcount = Workbooks(b).Sheets.Count - 1
                        NeueMappe.Range("G1").Value = plcount                      'Anzahl offener Projekte
                        NeueMappe.Cells(2, offsetSpalte).Value = Projektblatt.Name 'Projektnummer in kopfzeile anlegen
                        NeueMappe.Cells(ausgabezeile, offsetSpalte).Value = Projektblatt.Name 'Projektnummer in aktueller zeile anlegen
                        NeueMappe.Cells(ausgabezeile, offsetSpalte + 1).Value = Projektblatt.Cells(k, 8).Value - Projektblatt.Cells(k, 7).Value 'projektdifferenz setzen
                        NeueMappe.Cells(ausgabezeile, 7).Value = NeueMappe.Cells(ausgabezeile, 7).Value + NeueMappe.Cells(ausgabezeile, offsetSpalte + 1).Value 'differenz für diesen artikel aufsummieren
                        NeueMappe.Cells(ausgabezeile, offsetSpalte + 2).Value = Projektblatt.Cells(k, 9).Value 'Zu Wann übernehmen
                        'NeueMappe.Cells(ausgabezeile, 9).Value = Projektblatt.Cells(k, 15).Value 'Bestellt übernehmen 'nicht wiederholt neu setzen, bestellt-menge ist nicht projektspezifisch, sondern nur in der stammdatenliste hinterlegt
                        NeueMappe.Cells(ausgabezeile, offsetSpalte + 3).Value = Projektblatt.Cells(k, 10).Value 'Gemeldet am
                        NeueMappe.Cells(ausgabezeile, offsetSpalte + 4).Value = Projektblatt.Cells(k, 12).Value 'wer
                           ' If NeueMappe.Cells(ausgabezeile, 9).Value < NeueMappe.Cells(ausgabezeile, 7).Value Then 'bedingte Formatierung direkt in der vorlage bit excel bordmitteln
                            '    NeueMappe.Cells(ausgabezeile, 9).Interior.ColorIndex = 22
                           ' End If
                        If IsDate(NeueMappe.Cells(ausgabezeile, offsetSpalte + 3).Value) And IsDate(NeueMappe.Cells(ausgabezeile, 10).Value) Then
                            If CDate(NeueMappe.Cells(ausgabezeile, offsetSpalte + 3).Value) > CDate(NeueMappe.Cells(ausgabezeile, 10).Value) Then
                                NeueMappe.Cells(ausgabezeile, 10).Value = NeueMappe.Cells(ausgabezeile, offsetSpalte + 3).Value
                            End If
                        End If
                        
                   
                    End If
                Next k
            End If
            'End If
        Next i
        'bedarfsdifferenz auch aus lagerliste holen
        Set Projektblatt = Workbooks(b).Worksheets(1)
        If Projektblatt.UsedRange.Rows.Count > 1 Then
        
            For k = 2 To Projektblatt.UsedRange.Rows.Count
                If Projektblatt.Cells(k, 9).Value = "Nachbestellen" Then
                    'nix
                ElseIf Not IsNumeric(Projektblatt.Cells(k, 9).Value) Then
                    MsgBox "Lagerbestand in Zeile " & k & " kann nicht als Zahl interpretiert werden"
                    Exit For
                ElseIf Not IsNumeric(Projektblatt.Cells(k, 10).Value) Then
                    MsgBox "Lagerbedarf in Zeile " & k & " kann nicht als Zahl interpretiert werden"
                    Exit For
                ElseIf CSng(Projektblatt.Cells(k, 10).Value) > CSng(Projektblatt.Cells(k, 9).Value) Then
                    'MsgBox (Projektblatt.Cells(k, 9).Value & ";;" & Projektblatt.Cells(k, 10).Value)
                    If ersterTreffer Then       ' kopfzeile anlegen
                        ersterTreffer = False
                        Workbooks.Add (a + "\Report Bedarf.xltm")
                        Set NeueMappe = ActiveWorkbook.Sheets(1)
                        NeueMappe.Cells(1, 10).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
                        anzahlProjekte = 0
                    
                    End If
                    'gibts schon oder nicht?
                        
                    Set treffer = Nothing
                    Set treffer = NeueMappe.Columns(1).Find(what:=Projektblatt.Cells(k, 1).Value, lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
                    Set Spalte = Nothing    'spalte mit projekt finden
                    Set Spalte = NeueMappe.Rows(2).Find(what:="Lagerbedarf", lookat:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=False)
                    If Spalte Is Nothing Then
                        offsetSpalte = 11 + 5 * anzahlProjekte
                        anzahlProjekte = anzahlProjekte + 1
                    Else
                        offsetSpalte = Spalte.Column
                    End If
                    If treffer Is Nothing Then 'ean noch nicht angelegt
                        NeueMappe.Rows(3).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
                        Projektblatt.Range(Projektblatt.Cells(k, 1), Projektblatt.Cells(k, 6)).Copy Destination:=NeueMappe.Cells(3, 1) 'Artikeldaten aus projektmappe erstmals in report kopieren
                        NeueMappe.Cells(3, 7).Value = 0 'Summe auf 0 initialisieren
                        NeueMappe.Cells(3, 8).Value = Projektblatt.Cells(k, 9).Value   'istbestand initialisieren
                        NeueMappe.Cells(3, 9).Value = Projektblatt.Cells(k, 15).Value 'Bestellt einmalig auf den in der stammdatenliste hinterlegten wert setzen
                        NeueMappe.Cells(3, 10).Value = Projektblatt.Cells(k, 12).Value 'letzte Bedarfsmeldung für diesen artikel übernehmen
                        
                        ausgabezeile = 3
                       
                    Else 'ean schon da, addieren und projekt hinweis
                        ausgabezeile = treffer.Row
                    End If
               
                    NeueMappe.Cells(2, offsetSpalte).Value = "Lagerbedarf" 'Projektnummer in kopfzeile anlegen
                    NeueMappe.Cells(ausgabezeile, offsetSpalte).Value = "Lagerbedarf" 'Projektnummer in aktueller zeile anlegen
                    'NeueMappe.Cells(ausgabezeile, offsetSpalte + 1).Value = Projektblatt.Cells(k, 10).Value - Projektblatt.Cells(k, 9).Value 'projektdifferenz setzen
                    NeueMappe.Cells(ausgabezeile, offsetSpalte + 1).Value = Projektblatt.Cells(k, 10).Value 'projektbedarf setzen
                    NeueMappe.Cells(ausgabezeile, 7).Value = NeueMappe.Cells(ausgabezeile, 7).Value + Projektblatt.Cells(k, 10).Value - Projektblatt.Cells(k, 9).Value 'differenz für diesen artikel aufsummieren
                    NeueMappe.Cells(ausgabezeile, offsetSpalte + 2).Value = Projektblatt.Cells(k, 13).Value 'Zu Wann übernehmen
                    NeueMappe.Cells(ausgabezeile, offsetSpalte + 3).Value = Projektblatt.Cells(k, 12).Value 'Gemeldet am
                    NeueMappe.Cells(ausgabezeile, offsetSpalte + 4).Value = Projektblatt.Cells(k, 14).Value 'wer
                   '     If NeueMappe.Cells(ausgabezeile, 9).Value < NeueMappe.Cells(ausgabezeile, 7).Value Then 'bedingte Formatierung mit excel bordmitteln in der vorlage
                    '            NeueMappe.Cells(ausgabezeile, 9).Interior.ColorIndex = 22
                     '       End If
                    If IsDate(NeueMappe.Cells(ausgabezeile, offsetSpalte + 3).Value) And IsDate(NeueMappe.Cells(ausgabezeile, 10).Value) Then
                        If CDate(NeueMappe.Cells(ausgabezeile, offsetSpalte + 3).Value) > CDate(NeueMappe.Cells(ausgabezeile, 10).Value) Then
                            NeueMappe.Cells(ausgabezeile, 10).Value = NeueMappe.Cells(ausgabezeile, offsetSpalte + 3).Value
                        End If
                    End If
                End If
            Next k
        End If
        'End If
            
        'Nutzerkürzel als Hilfszelle in Report überführen für automatischen Bestell-BatchBuchung
        If fixeduser <> "??" Then
            ActiveWorkbook.Worksheets(2).Cells(5, 5).Value = fixeduser
        Else
            ActiveWorkbook.Worksheets(2).Cells(5, 5).Value = Workbooks(Dateiname).Worksheets(3).Cells(2, 1).Value
        End If
        'Dteiname terminal als Hilfszelle in Report überführen für automatischen Bestell-BatchBuchung
        ActiveWorkbook.Worksheets(2).Cells(50, 1).Value = Dateiname
        ActiveWorkbook.Worksheets(1).Activate
        
        
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
           ' Beep 2000, 750
        End If
        NeueMappe.Rows("1:100").RowHeight = 15 'workaround gegen automatische formatierungen
        
    Else
        MsgBox "keine Projektblätter gefunden"
    End If
    Workbooks(b).Close SaveChanges:=False
Exit Sub
    
FehlerOpen:
    MsgBox "Fehler in Sub ReportBedarf beim Öffnen" & vbCrLf & Err.Number & vbCrLf & Err.Description

Exit Sub

FehlerNachOpen:
    MsgBox "Fehler in Sub ReportBedarf nach Öffnen" & vbCrLf & Err.Number & vbCrLf & Err.Description
    On Error Resume Next
    Workbooks(b).Close SaveChanges:=False
    
End Sub

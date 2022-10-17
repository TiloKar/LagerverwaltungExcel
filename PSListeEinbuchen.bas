Attribute VB_Name = "PSListeEinbuchen"
Public Sub BestandAnhängen()
    Call holeDatenbank
    Dateiname = ActiveWorkbook.Name
    Dim terminal As Worksheet
    Set terminal = Workbooks(Dateiname).Worksheets(1)
    'hier bestandsbuchungen nach vollständiger abarbeitung einmal anhängen
    For i = terminal.UsedRange.Cells(terminal.UsedRange.Rows.Count, 1).Row To 2 Step -1
        Dim treffer As Integer
        If terminal.Cells(i, 1).Value = "Bedarf" Then
            treffer = 0
            For k = 1 To llrows
                If StrComp(terminal.Cells(i, 4), ll(k, 1), 1) = 0 Then
                     treffer = k
                     Exit For
                End If
            Next k
             
             If treffer > 0 Then
                 If IsNumeric(ll(treffer, 9)) Then
                    If ll(treffer, 9) > 0 Then
                        'neu einfügen
                        terminal.Rows(i).Insert Shift:=xlShiftDown
                        terminal.Range(terminal.Cells(i + 1, 1), terminal.Cells(i + 1, 9)).Copy Destination:=terminal.Cells(i, 1)
                        terminal.Cells(i, 1).Value = "Bestand"
                        If ll(treffer, 9) < terminal.Cells(i + 1, 3) Then
                            terminal.Cells(i, 3).Value = ll(treffer, 9)
                         Else
                            terminal.Cells(i, 3).Value = terminal.Cells(i + 1, 3)
                        End If
                      
                    End If
                End If
            End If
        End If
    Next i
End Sub

Public Sub BedarfPSImport()
    If BatchBuchungen.Projektauswahl.ListIndex = -1 Then
        MsgBox "Projekt wählen"
    ElseIf BatchBuchungen.Nutzer.ListIndex = -1 Then
        MsgBox "Nutzer wählen"
    Else
        
        strListe = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx, Excel 97 Files (*.xls), *.xls")
        If strListe = False Then
            MsgBox "Abgebrochen"
        Else
            keinPiep = True
            Dim terminal As Worksheet
            Set terminal = Workbooks(Dateiname).Worksheets(1)
            Dim stklisteDatei As Workbook
           ' On Error GoTo FehlerOpen
            Set stklisteDatei = Workbooks.Open(Filename:=strListe, ReadOnly:=True)
            On Error GoTo 0
            Dim stkliste As Worksheet
            Set stkliste = stklisteDatei.Worksheets(1)
            
            Dim NeueMappe As Worksheet
            Workbooks.Add (a + "\Checkliste Roter Punkt.xltx") 'mappe anlegen für roten punkt
            Set NeueMappe = ActiveWorkbook.Sheets(1)
            NeueMappe.Cells(1, 1).Value = "Checkliste Roter Punkt, " + BatchBuchungen.Projektauswahl.Text + " für: "
            NeueMappe.Cells(1, 9).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
            NeueMappe.Cells(2, 1).Value = strListe

            rpZeile = 3
         
            Dim fund1 As Integer
            Dim fund2 As Integer
            If llrows > 1 And llrows < 10000 And stkliste.UsedRange.Cells(1, stkliste.UsedRange.Columns.Count).Column >= 9 Then 'Stkliste.UsedRange.Columns.Count >= 9 Then
                Dim treffer1 As Integer
                Dim treffer2 As Integer
                For k = 9 To stkliste.UsedRange.Cells(1, stkliste.UsedRange.Columns.Count).Column
                    fund1 = 1
                    fund2 = 1
                    If stkliste.Cells(1, k).Value = "" Or stkliste.Cells(3, k).Value = 0 Or stkliste.Cells(3, k).Value = "" Then
                      'nix wenn leer oder 0
                    Else
                        treffer1 = 0
                        For i = 2 To llrows
                            If StrComp(stkliste.Cells(1, k).Value, ll(i, 1), 1) = 0 Then
                                 treffer1 = i
                                 Exit For
                            End If
                        Next i
                        
                        If treffer1 = 0 Then
                             'bei keinem treffer spalte 2 und 3 kopieren und hinweisspalte in buchungsliste beschreiben
                            terminal.Rows(2).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
                            terminal.Rows(2).Font.Color = RGB(255, 0, 0)
                            terminal.Cells(2, 1).Value = "Bedarf" 'buchungsart
                            terminal.Cells(2, 2).Value = BatchBuchungen.Projektauswahl.Text 'projekt
                            terminal.Cells(2, 3).Value = stkliste.Cells(3, k).Value 'stückzahl
                            terminal.Cells(2, 4).Value = stkliste.Cells(1, k).Value 'scancode
                            terminal.Cells(2, 5).Value = BatchBuchungen.Wann.Value  'wann
                            terminal.Cells(2, 6).Value = BatchBuchungen.Nutzer.Text ' wer
                            terminal.Cells(2, 7).Value = stkliste.Cells(2, k).Value '2.zeile PS liste als hinweis in kurzbez 1
                            terminal.Cells(2, 10).Value = "kein Treffer"
                        Else
                            treffer2 = 0
                            For i = treffer1 To llrows
                                If StrComp(stkliste.Cells(1, k).Value, ll(i, 1), 1) = 0 Then
                                     treffer2 = i
                                     Exit For
                                End If
                            Next i
                            If treffer2 > treffer1 Then   'mehrfachtreffer
                                terminal.Rows(2).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
                                terminal.Rows(2).Font.Color = RGB(255, 0, 0)
                                terminal.Cells(2, 1).Value = "Bedarf" 'buchungsart
                                terminal.Cells(2, 2).Value = BatchBuchungen.Projektauswahl.Text 'projekt
                                terminal.Cells(2, 3).Value = stkliste.Cells(3, k).Value 'stückzahl
                                terminal.Cells(2, 4).Value = stkliste.Cells(1, k).Value 'scancode
                                terminal.Cells(2, 5).Value = BatchBuchungen.Wann.Value  'wann
                                terminal.Cells(2, 6).Value = BatchBuchungen.Nutzer.Text ' wer
                                terminal.Cells(2, 7).Value = stkliste.Cells(2, k).Value '2.zeile PS liste als hinweis in kurzbez 1
                                terminal.Cells(2, 10).Value = "!!! mehrfacher Treffer !!!"
                            Else        'gültiger einzeltreffer
                             
                                If ll(treffer1, 8) = "nein" Or ll(treffer1, 8) = "Nein" Then  'roter punkt
                                    'rote punkt artikel in mappe stapeln (merkliste)
                                    NeueMappe.Rows(3).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
                                    NeueMappe.Cells(3, 1).Value = stkliste.Cells(3, k).Value 'stückzahl
                                    NeueMappe.Cells(3, 2).Value = ll(treffer1, 1)
                                    NeueMappe.Cells(3, 3).Value = ll(treffer1, 2)
                                    NeueMappe.Cells(3, 4).Value = ll(treffer1, 3)
                                    NeueMappe.Cells(3, 5).Value = ll(treffer1, 4)
                                    rpZeile = rpZeile + 1
                                   
                                Else ' wenn ja bedarfzahl aus treffersplate 1 mit artikelnummer und bezeichner 1 und 2  mit projektnummer und wann-feld in buchungsliste anlegen
                                    Call BuchungInListeAnlegen("Bedarf", BatchBuchungen.Projektauswahl.Text, stkliste.Cells(3, k).Value, ll(treffer1, 1), BatchBuchungen.Wann.Value, ll(treffer1, 2), ll(treffer1, 3), ll(treffer1, 4), BatchBuchungen.Nutzer.Text)
                                    
                                End If
                            End If
                        End If
                    End If
                Next k
             End If
             
            '  Call BestandAnhängen(Workbooks(Dateiname).Worksheets(1))
             
            stklisteDatei.Close SaveChanges:=False
            terminal.Rows("2:200").RowHeight = 15
            BatchBuchungen.Hide
            keinPiep = False
        End If
        
    End If


FehlerOpen:
    
        
End Sub


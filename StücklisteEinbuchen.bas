Attribute VB_Name = "StücklisteEinbuchen"


Public Sub BedarfStücklisteImport(CAD As Boolean)
    Call holeDatenbank
    If BatchBuchungen.Projektauswahl.ListIndex = -1 Then
        MsgBox "Projekt wählen"
    ElseIf BatchBuchungen.Nutzer.ListIndex = -1 Then
        MsgBox "Nutzer wählen"
    Else
        Dim startspalte As Integer
        If CAD Then
            startspalte = 2
        Else
            startspalte = 1
        End If
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
            Dim fs As Object
            Set fs = CreateObject("Scripting.FileSystemObject")
            On Error Resume Next
            If CAD Then     'ordner für pdf/step erzeugen
                Dim pathFolder As String
                pathFolder = stklisteDatei.Path
                pathFolder = pathFolder + "\Dateien_" + stklisteDatei.Name
                pathFolder = Left(pathFolder, (InStrRev(pathFolder, ".") - 1)) 'endung abschneiden
                'If Dir(pathFolder, vbDirectory) = "" Then
                '    MkDir (pathFolder)
                'End If
                If fs.FolderExists(pathFolder) = False Then
                    fs.CreateFolder (pathFolder)
                End If
                
            End If

                
            Dim NeueMappe As Worksheet
            Workbooks.Add (a + "\Checkliste Roter Punkt.xltx") 'mappe anlegen für roten punkt
            Set NeueMappe = ActiveWorkbook.Sheets(1)
            NeueMappe.Cells(1, 1).Value = "Checkliste Roter Punkt, " + BatchBuchungen.Projektauswahl.Text + " für: "
            NeueMappe.Cells(1, 9).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
            NeueMappe.Cells(2, 1).Value = strListe

            rpZeile = 3
       
            Dim fund1 As Integer
            Dim fund2 As Integer
            Dim srcFileString As String
            Dim destFileString As String
            Dim crlfString As String
            Dim copyOK As Boolean
            If llrows > 1 And llrows < 10000 Then
                Dim treffer1 As Integer
                Dim treffer2 As Integer
                'MsgBox (Stkliste.UsedRange.Cells(Stkliste.UsedRange.Rows.Count, 1).Row)
                
                For k = 3 To stkliste.UsedRange.Cells(stkliste.UsedRange.Rows.Count, 1).Row
                    
                    
                    fund1 = 1
                    fund2 = 1
                    
                    If stkliste.Cells(k, startspalte + 1).Value = "" Or stkliste.Cells(k, startspalte).Value = 0 Or stkliste.Cells(k, startspalte).Value = "" Then
                      'nix wenn leer oder 0
                    Else
                 
                        
                        treffer1 = 0
                        For i = 2 To llrows
                            If StrComp(stkliste.Cells(k, startspalte + 1).Value, ll(i, 1), 1) = 0 Then
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
                            terminal.Cells(2, 3).Value = stkliste.Cells(k, startspalte).Value 'stückzahl
                            
                            crlfString = stkliste.Cells(k, startspalte + 1).Value 'scancode um crlf bereinigen
                            crlfString = Replace(crlfString, vbCrLf, "", 1, -1, 1)
                            terminal.Cells(2, 4).Value = crlfString
                            
                            terminal.Cells(2, 5).Value = BatchBuchungen.Wann.Value  'wann
                            terminal.Cells(2, 6).Value = BatchBuchungen.Nutzer.Text ' wer
                            terminal.Cells(2, 7).Value = stkliste.Cells(k, startspalte + 2).Value '3.Spalte stückliste als hinweis in kurzbez 1
                            terminal.Cells(2, 8).Value = stkliste.Cells(k, startspalte + 3).Value '3.Spalte stückliste als hinweis in kurzbez 1
                            terminal.Cells(2, 9).Value = stkliste.Cells(k, startspalte + 4).Value '3.Spalte stückliste als hinweis in kurzbez 1
                            terminal.Cells(2, 10).Value = "kein Treffer"
                            
                            'pdf/step kopieren
                            On Error Resume Next
                            If CAD And (stkliste.Cells(k, 5).Value <> "") And (stkliste.Cells(k, 5).Value <> "-") Then
                                srcFileString = stkliste.Cells(k, 12).Value & stkliste.Cells(k, 11).Value & ".pdf"
                                srcFileString = Replace(srcFileString, vbCrLf, "", 1, -1, 1)
                                destFileString = pathFolder & "/" & stkliste.Cells(k, 11).Value & ".pdf"
                                destFileString = Replace(destFileString, vbCrLf, "", 1, -1, 1)
                                
                                copyOK = False
                                If fs.FileExists(srcFileString) Then
                                    fs.CopyFile srcFileString, destFileString, True
                                    copyOK = fs.FileExists(destFileString)
                                End If
                                If copyOK Then
                                    terminal.Cells(2, 11).Value = "*.pdf kopiert!"
                                Else
                                    terminal.Cells(2, 11).Value = "*.pdf nicht gefunden!"
                                End If
                                
                                srcFileString = stkliste.Cells(k, 12).Value & stkliste.Cells(k, 11).Value & ".step"
                                destFileString = pathFolder & "/" & stkliste.Cells(k, 11).Value & ".step"
                                
                                copyOK = False
                                If fs.FileExists(srcFileString) Then
                                    fs.CopyFile srcFileString, destFileString, True
                                    copyOK = fs.FileExists(destFileString)
                                End If
                                If copyOK Then
                                    terminal.Cells(2, 11).Value = terminal.Cells(2, 11).Value & ", *.step kopiert!"
                                End If
                                
                            End If
                            On Error GoTo 0
                            
                            
                            
                        Else
                            treffer2 = 0
                            For i = treffer1 To llrows
                                If StrComp(stkliste.Cells(k, startspalte + 1).Value, ll(i, 1), 1) = 0 Then
                                     treffer2 = i
                                     Exit For
                                End If
                            Next i
                            If treffer2 > treffer1 Then   'mehrfachtreffer
                                terminal.Rows(2).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
                                terminal.Rows(2).Font.Color = RGB(255, 0, 0)
                                terminal.Cells(2, 1).Value = "Bedarf" 'buchungsart
                                terminal.Cells(2, 2).Value = BatchBuchungen.Projektauswahl.Text 'projekt
                                terminal.Cells(2, 3).Value = stkliste.Cells(k, startspalte).Value 'stückzahl
                                terminal.Cells(2, 4).Value = stkliste.Cells(k, startspalte + 1).Value 'scancode
                                terminal.Cells(2, 5).Value = BatchBuchungen.Wann.Value  'wann
                                terminal.Cells(2, 6).Value = BatchBuchungen.Nutzer.Text ' wer
                                terminal.Cells(2, 7).Value = stkliste.Cells(k, startspalte + 2).Value '3.Spalte stückliste als hinweis in kurzbez 1
                                terminal.Cells(2, 8).Value = stkliste.Cells(k, startspalte + 3).Value '3.Spalte stückliste als hinweis in kurzbez 1
                                terminal.Cells(2, 9).Value = stkliste.Cells(k, startspalte + 4).Value '3.Spalte stückliste als hinweis in kurzbez 1
                                terminal.Cells(2, 10).Value = "!!! mehrfacher Treffer !!!"
                            Else        'gültiger einzeltreffer
                                If ll(treffer1, 8) = "nein" Or ll(treffer1, 8) = "Nein" Then  'roter punkt
                                    'rote punkt artikel in mappe stapeln (merkliste)
                                    NeueMappe.Rows(3).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
                                    NeueMappe.Cells(3, 1).Value = stkliste.Cells(k, startspalte).Value 'stückzahl
                                    NeueMappe.Cells(3, 2).Value = ll(treffer1, 1)
                                    NeueMappe.Cells(3, 3).Value = ll(treffer1, 2)
                                    
                                    NeueMappe.Cells(3, 4).Value = ll(treffer1, 3)
                                    NeueMappe.Cells(3, 5).Value = ll(treffer1, 4)
                                    rpZeile = rpZeile + 1
                                   
                                Else ' wenn ja bedarfzahl aus treffersplate 1 mit artikelnummer und bezeichner 1 und 2  mit projektnummer und wann-feld in buchungsliste anlegen
                                    Call BuchungInListeAnlegen("Bedarf", BatchBuchungen.Projektauswahl.Text, stkliste.Cells(k, startspalte).Value, ll(treffer1, 1), BatchBuchungen.Wann.Value, ll(treffer1, 2), ll(treffer1, 3), ll(treffer1, 4), BatchBuchungen.Nutzer.Text)
                                    'pdf/step kopieren
                                    On Error Resume Next
                                    If CAD And (stkliste.Cells(k, 5).Value <> "") And (stkliste.Cells(k, 5).Value <> "-") Then
                                        srcFileString = stkliste.Cells(k, 12).Value & stkliste.Cells(k, 11).Value & ".pdf"
                                        srcFileString = Replace(srcFileString, vbCrLf, "", 1, -1, 1)
                                        destFileString = pathFolder & "/" & stkliste.Cells(k, 11).Value & ".pdf"
                                        destFileString = Replace(destFileString, vbCrLf, "", 1, -1, 1)
                                        
                                        copyOK = False
                                        If fs.FileExists(srcFileString) Then
                                            fs.CopyFile srcFileString, destFileString, True
                                            copyOK = fs.FileExists(destFileString)
                                        End If
                                        If copyOK Then
                                            terminal.Cells(2, 11).Value = "*.pdf kopiert!"
                                        Else
                                            terminal.Cells(2, 11).Value = "*.pdf nicht gefunden!"
                                        End If
                                        
                                        srcFileString = stkliste.Cells(k, 12).Value & stkliste.Cells(k, 11).Value & ".step"
                                        destFileString = pathFolder & "/" & stkliste.Cells(k, 11).Value & ".step"
                                        
                                        copyOK = False
                                        If fs.FileExists(srcFileString) Then
                                            fs.CopyFile srcFileString, destFileString, True
                                            copyOK = fs.FileExists(destFileString)
                                        End If
                                        If copyOK Then
                                            terminal.Cells(2, 11).Value = terminal.Cells(2, 11).Value & ", *.step kopiert!"
                                        End If
                                        
                                    End If
                                    On Error GoTo 0
                                End If
                            End If
                        End If
                    End If
                Next k
             End If
                
            'Call BestandAnhängen(Workbooks(Dateiname).Worksheets(1))
                
            stklisteDatei.Close SaveChanges:=False
            terminal.Rows("2:200").RowHeight = 15
            BatchBuchungen.Hide
            keinPiep = False
        End If
        
    End If
    
FehlerOpen:
Exit Sub

    
End Sub


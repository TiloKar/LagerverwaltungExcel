Attribute VB_Name = "SLPEinbuchen"
Private Sub findeSLPTeil(ByRef terminal As Worksheet, ByRef NeueMappe As Worksheet, searchpart, searchvalue, anzahl)
 'finde in lagerliste, dann buchen oder rp
'MsgBox ("vergleichFUN: |" & searchpart & "|" & searchvalue & "|")
    Dim treffer1 As Integer
    Dim treffer2 As Integer
    Dim trefferTerminal As Integer
    Dim bestandTerminal As Integer
    treffer1 = 0
    For i = 2 To llrows
'        If searchvalue <> "" And StrComp(searchvalue, ll(i, 6), 1) = 0 Then
'            MsgBox ("vergleichFUN: |" & searchpart & "|" & ll(i, 5) & "|----|" & searchvalue & "|" & ll(i, 6) & "|")
'        End If
        If StrComp(searchpart, ll(i, 5), 1) = 0 And StrComp(searchvalue, ll(i, 6), 1) = 0 Then
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
        terminal.Cells(2, 3).Value = anzahl 'stückzahl
        terminal.Cells(2, 4).Value = "?" 'scancode
        terminal.Cells(2, 5).Value = BatchBuchungen.Wann.Value  'wann
        terminal.Cells(2, 6).Value = BatchBuchungen.Nutzer.Text ' wer
        terminal.Cells(2, 7).Value = searchpart '
        terminal.Cells(2, 8).Value = searchvalue '
        terminal.Cells(2, 10).Value = "kein Treffer"
    Else
        treffer2 = 0
        For i = treffer1 To llrows
            If StrComp(searchpart, ll(i, 5), 1) = 0 And StrComp(searchvalue, ll(i, 6), 1) = 0 Then
                 treffer2 = i
                 Exit For
            End If
        Next i
        If treffer2 > treffer1 Then   'mehrfachtreffer
            terminal.Rows(2).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
            terminal.Rows(2).Font.Color = RGB(255, 0, 0)
            terminal.Cells(2, 1).Value = "Bedarf" 'buchungsart
            terminal.Cells(2, 2).Value = BatchBuchungen.Projektauswahl.Text 'projekt
            terminal.Cells(2, 3).Value = anzahl 'stückzahl
            terminal.Cells(2, 4).Value = ll(treffer2, 1) 'scancode treffer 2
            terminal.Cells(2, 5).Value = BatchBuchungen.Wann.Value  'wann
            terminal.Cells(2, 7).Value = searchpart '
            terminal.Cells(2, 8).Value = searchvalue '
            terminal.Cells(2, 10).Value = "!!! mehrfacher Treffer !!!"
        Else        'gültiger einzeltreffer

            If ll(treffer1, 8) = "nein" Or ll(treffer1, 8) = "Nein" Then  'roter punkt
                'rote punkt artikel in mappe stapeln (merkliste)
                NeueMappe.Rows(3).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
                NeueMappe.Cells(3, 1).Value = anzahl 'stückzahl
                NeueMappe.Cells(3, 2).Value = ll(treffer1, 1)
                NeueMappe.Cells(3, 3).Value = ll(treffer1, 2)
                NeueMappe.Cells(3, 4).Value = ll(treffer1, 3)
                NeueMappe.Cells(3, 5).Value = ll(treffer1, 4)
                rpZeile = rpZeile + 1

            Else ' wenn ja bedarfzahl aus treffersplate 1 mit artikelnummer und bezeichner 1 und 2  mit projektnummer und wann-feld in buchungsliste anlegen
                Call BuchungInListeAnlegen("Bedarf", BatchBuchungen.Projektauswahl.Text, anzahl, ll(treffer1, 1), BatchBuchungen.Wann.Value, ll(treffer1, 2), ll(treffer1, 3), ll(treffer1, 4), BatchBuchungen.Nutzer.Text)
                'falls lagerbestand vorhanden dann passende zeile anlegen
'                If IsNumeric(ll(treffer1, 9)) Then
'                    If ll(treffer1, 9) > 0 Then
'                        If anzahl > 0 Then
'                            If ll(treffer1, 9) < anzahl - bestandTerminal Then
'                                Call BuchungInListeAnlegen("Bestand", BatchBuchungen.Projektauswahl.Text, ll(treffer1, 9), ll(treffer1, 1), BatchBuchungen.Wann.Value, ll(treffer1, 2), ll(treffer1, 3), ll(treffer1, 4), BatchBuchungen.Nutzer.Text)
'                            Else
'                                Call BuchungInListeAnlegen("Bestand", BatchBuchungen.Projektauswahl.Text, anzahl, ll(treffer1, 1), BatchBuchungen.Wann.Value, ll(treffer1, 2), ll(treffer1, 3), ll(treffer1, 4), BatchBuchungen.Nutzer.Text)
'                            End If
'                        End If
'                    End If
'                End If
            End If
        End If
    End If
End Sub


Public Sub BedarfSLPImport()
    Call holeDatenbank
    If BatchBuchungen.Projektauswahl.ListIndex = -1 Then
        MsgBox "Projekt wählen"
    ElseIf BatchBuchungen.Nutzer.ListIndex = -1 Then
        MsgBox "Nutzer wählen"
    Else
        
        strListe = Application.GetOpenFilename("Text Files (*.txt), *.txt")
        If strListe = False Then
            MsgBox "Abgebrochen"
        Else
            Dim NeueMappe As Worksheet
            Workbooks.Add (a + "\Checkliste Roter Punkt.xltx") 'mappe anlegen für roten punkt
            Set NeueMappe = ActiveWorkbook.Sheets(1)
            NeueMappe.Cells(1, 1).Value = "Checkliste Roter Punkt, " + BatchBuchungen.Projektauswahl.Text + " für: "
            NeueMappe.Cells(1, 9).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
            NeueMappe.Cells(2, 1).Value = strListe
            rpZeile = 3
            keinPiep = True
            
            stklisteDatei = FreeFile
            Open strListe For Binary As stklisteDatei
            Dim strStkliste As String
            strStkliste = Space(LOF(stklisteDatei))
            Get stklisteDatei, 1, strStkliste
            Close stklisteDatei
            
            Dim pos1 As Long
            pos1 = 1
            Dim pos2 As Long
            Dim pos2alt As Long
            pos2 = 1
            pos2alt = 1
            Dim hlpstr As String
            For i = 1 To 5
                pos1 = InStr(pos1, strStkliste, "|" & vbCrLf, vbBinaryCompare) + 3  'pos1 auf zeile 1 setzen
            Next i
            pos1 = pos1 + 1 ' estes | der zeile überspringen
            
            
            Dim nrow As Integer
            nrow = 1
            Dim ncol As Integer
            ncol = 1
            Dim anzahl(maxRowsSLP) As Integer
            Dim partname(maxRowsSLP) As String
            Dim valprop(maxRowsSLP) As String
            Dim searchpart As String
            Dim searchvalue As String
            Dim pos As Long
            Dim lastpart As String
            Dim valuegiven As Boolean
            
            While pos2 + 3 < Len(strStkliste)
                pos2 = InStr(pos1, strStkliste, "|", vbTextCompare)             'pos2 auf nächstes | setzen
                pos2alt = InStr(pos1, strStkliste, "|" & vbCrLf, vbTextCompare) 'pos2alt auf nächstes | und CRLF setzen
                hlpstr = Mid(strStkliste, pos1, pos2 - pos1)
                If ncol = 1 Then
                    If Not IsNumeric(hlpstr) Then
                        pos2 = Len(strStkliste) ' abbruch falls spalte 1 keine zahl mehr enthält
                        pos2alt = pos2 + 1
                        posFault = 0
                        posFault = InStr(1, hlpstr, "-", vbTextCompare) 'fallunterscheidung ob Formatfehler oder letzte zeile
                        If posFault = 0 Then
                            MsgBox ("Unerwarteter Zeilenumbruch. Bitte Buchungsaufträge verwerfen und Textdatei prüfen. Nach manueller Korrektur bitte neu einbuchen ")
                        End If
                    Else
                        anzahl(nrow) = Int(hlpstr)
                    End If
                ElseIf ncol = 2 Then
                    partname(nrow) = RTrim(hlpstr)
                    pos2part = Len(partname(nrow))
                    pos1part = 1
                    lastpart = ""
                    valuegiven = False
                ElseIf ncol = 3 Then
                    valprop(nrow) = RTrim(hlpstr)
                End If
                            
                'value niemals mit +, bezieht sich auf letzgenanntes partname
                'wenn letzgenannte partnames gleich, dann auch value auf alle gleichen anwenden
                If pos2 = pos2alt Then 'falls endzeichen mit zeilenumbruch = koplette zeile geparst
                    If partname(nrow) <> "" And partname(nrow) <> "ignore" Then
                        pos = InStrRev(partname(nrow), " + ")   'mehrfachbelegung erkennen
                        
                        While pos > 0
                            
                            searchpart = Right(partname(nrow), Len(partname(nrow)) - pos - 2)   'searchpart beschreiben
                            partname(nrow) = Left(partname(nrow), Len(partname(nrow)) - Len(searchpart) - 3) 'gesamtstring um extrahierten part und um + zeichen bereinigen
                           
                            If Not valuegiven Then
                                If lastpart = "" Then   ' beim ersten fund von hinten value einfach vergeben
                                   searchvalue = valprop(nrow)
                                Else
                                    If lastpart = searchpart Then   ' solange gleiches value vergeben, bis part name von hinten sich ändert
                                        searchvalue = valprop(nrow)
                                    Else
                                        valuegiven = True   ' dann kein value mehr
                                        searchvalue = ""
                                        valprop(nrow) = ""
                                    End If
                                End If
                                lastpart = searchpart
                            End If
'                            MsgBox ("vergleichRUF1: |" & searchpart & "|" & searchvalue & "|")
                            Call findeSLPTeil(Workbooks(Dateiname).Worksheets(1), NeueMappe, searchpart, searchvalue, anzahl(nrow))
                            
                            pos = InStrRev(partname(nrow), " + ")
                        Wend
                        'erstgenannter part
                       
                       searchpart = partname(nrow)
                       If Not valuegiven Then
                            If lastpart = "" Then   ' beim ersten fund von hinten value einfach vergeben
                               searchvalue = valprop(nrow)
                            Else
                                If lastpart = searchpart Then   ' solange gleiches value vergeben, bis part name von hinten sich ändert
                                    searchvalue = valprop(nrow)
                                Else
                                    valuegiven = True   ' dann kein value mehr
                                    searchvalue = ""
                                    valprop(nrow) = ""
                                End If
                            End If
                            lastpart = searchpart
                        End If
'                        MsgBox ("vergleichRUF2: |" & searchpart & "|" & searchvalue & "|")
                        Call findeSLPTeil(Workbooks(Dateiname).Worksheets(1), NeueMappe, searchpart, searchvalue, anzahl(nrow))
                        
                    End If
                    
                    pos1 = pos2 + 4
                    ncol = 1
                    nrow = nrow + 1
                    
                Else
                    pos1 = pos2 + 1
                    ncol = ncol + 1
                End If
  
            Wend
            
            'Call BestandAnhängen(Workbooks(Dateiname).Worksheets(1))

            Workbooks(Dateiname).Worksheets(1).Rows("2:200").RowHeight = 15
            BatchBuchungen.Hide
            keinPiep = False
        End If
        
    End If
    
FehlerOpen:
 
End Sub


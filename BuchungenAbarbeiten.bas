Attribute VB_Name = "BuchungenAbarbeiten"
Public Sub journalEintragAnlegen(ByRef journal As Worksheet, ByVal zeile As Integer)
    Dim auftr�ge As Worksheet        'tempor�re arbeitsblatt
    Set auftr�ge = Workbooks(b).Worksheets(1)
    journal.Rows(1).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    journal.Cells(1, 1).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss") 'Zeit
    journal.Cells(1, 2).Value = auftr�ge.Cells(zeile, 1).Value ' EAN
    journal.Cells(1, 3).Value = auftr�ge.Cells(zeile, 2).Value ' Bez 1
    
    
    
End Sub





Public Sub Problemkinder(ByRef auftr�ge As Worksheet, ByRef journal As Worksheet, Index As Integer)

    On Error GoTo FehlerOpenProblemkinder
    Workbooks.Open Filename:=a & d, ReadOnly:=False
    On Error GoTo FehlerProblemkinder
    journal.Rows(1).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    journal.Cells(1, 1).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss") ' Zeit
    journal.Cells(1, 2).Value = auftr�ge.Cells(Index, 4).Value ' EAN
    journal.Cells(1, 3).Value = auftr�ge.Cells(Index, 7).Value ' Bez 1
    journal.Cells(1, 4).Value = auftr�ge.Cells(Index, 10).Value ' Fehlertext
    journal.Cells(1, 5).Value = "Zu Problemkinder verschoben"
    journal.Cells(1, 7).Value = auftr�ge.Cells(Index, 6).Value 'wer
    Workbooks(d).Sheets(1).Rows(2).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    auftr�ge.Range(auftr�ge.Cells(Index, 1), auftr�ge.Cells(Index, 10)).Cut Destination:=Workbooks(d).Sheets(1).Cells(2, 1)
    auftr�ge.Rows(Index).Delete
    Workbooks(d).Close SaveChanges:=True
Exit Sub
FehlerOpenProblemkinder:
    MsgBox "Fehlerhafter Buchungsauftrag in Zeile " & Index & " konnte nicht zu den M�dels �bermittelt werden und verbleibt bis zum n�chsten Versuch in dieser Liste"
Exit Sub
FehlerProblemkinder:
    On Error Resume Next
    MsgBox "Fehler bei der �bermittlung des fehlerhaften Buchungsauftrags in Zeile " & i & " zu den M�dels."
    Workbooks(d).Close SaveChanges:=True
End Sub



Public Sub ListeAbarbeiten()
       
    If Workbooks(Dateiname).Worksheets(3).Cells(11, 4).Value = "x" Then
        MsgBox "Im Offline Modus nicht m�glich!"
        Exit Sub
    End If
       
   'Call Inits
   Call holeDatenbank
   
    Dim auftr�ge As Worksheet        'tempor�re arbeitsblatt
    Set auftr�ge = Workbooks(Dateiname).Worksheets(1)
    Dim treffer As Object

    
    If auftr�ge.UsedRange.Rows.Count > 1 Then
        On Error GoTo FehlerOpenLagerliste
        Workbooks.Open Filename:=a & b, ReadOnly:=False, WriteResPassword:=pwlager, Password:=pwlager
        If Workbooks(b).ReadOnly Then
            GoTo LagerlisteSchonOffen
        End If
        On Error GoTo FehlerOpenJournal
        Dim lagerliste As Worksheet 'tempor�re arbeitsblatt
        Set lagerliste = Workbooks(b).Worksheets(1)
        Workbooks.Open Filename:=a & c, ReadOnly:=False, WriteResPassword:=pwjournal, Password:=pwjournal
        If Workbooks(c).ReadOnly Then
            GoTo JournalSchonOffen
        End If
        On Error GoTo FehlerVorBuchung
        Dim journal As Worksheet                'tempor�re arbeitsblatt
        Set journal = Workbooks(c).Worksheets(1)
        
        Dim i As Integer
        For i = auftr�ge.UsedRange.Rows.Count To 2 Step -1
     
            'Versuche EAN in Lagerliste zu finden f�r alle Buchungen
            Set treffer = Nothing
            Set treffer = lagerliste.Columns(1).Find(what:=auftr�ge.Cells(i, 4).Value, lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)

            Dim userindex As Integer
            userindex = -1
            For k = 0 To usercount
                If usernames(k) = auftr�ge.Cells(i, 6).Value Then
                    userindex = k
                    Exit For
                End If
            Next
        
            If auftr�ge.Cells(i, 1).Value = "" And auftr�ge.Cells(i, 2).Value = "" And auftr�ge.Cells(i, 3).Value = "" And auftr�ge.Cells(i, 4).Value = "" And auftr�ge.Cells(i, 6).Value = "" Then
                auftr�ge.Rows(i).Delete 'delete
                GoTo n�chsteZeile
            ElseIf auftr�ge.Cells(i, 1).Value = "" Or auftr�ge.Cells(i, 2).Value = "" Or auftr�ge.Cells(i, 3).Value = "" Or auftr�ge.Cells(i, 4).Value = "" Then
                auftr�ge.Cells(i, 10).Value = "Auftrag unvollst�ndig"
                GoTo n�chsteZeile
            ElseIf userindex = -1 Then
                auftr�ge.Cells(i, 10).Value = "unbekannter Nutzer"
                Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                GoTo n�chsteZeile
            ElseIf treffer Is Nothing Then
                auftr�ge.Cells(i, 10).Value = "Scancode in Lagerliste nicht gefunden"
                Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                GoTo n�chsteZeile
            ElseIf Not IsNumeric(lagerliste.Cells(treffer.Row, 9).Value) And lagerliste.Cells(treffer.Row, 9).Value <> "Nachbestellen" Then
                auftr�ge.Cells(i, 10).Value = "Bestand in Lagerliste l�sst sich nicht als Zahl interpretieren"
                Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                GoTo n�chsteZeile
            Else
                Dim TrefferLagerliste As Integer
                TrefferLagerliste = treffer.Row
            End If
            
            'Liste abarbeiten
            If auftr�ge.Cells(i, 1).Value = "Bestand" And auftr�ge.Cells(i, 2).Value = "Lager" And auftr�ge.Cells(i, 3).Value = "Nachbestellen" Then    'Bestellmarkierung
                 'oben einf�gen und neu anlegen
                 Call journalEintragAnlegen(journal, TrefferLagerliste)
                journal.Cells(1, 4).Value = "Nachbestellen markiert"
                journal.Cells(1, 7).Value = auftr�ge.Cells(i, 6).Value 'wer
                lagerliste.Cells(TrefferLagerliste, 9).Value = "Nachbestellen"
                lagerliste.Cells(TrefferLagerliste, 12).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss") 'letzte bedarfsmeldung
                lagerliste.Cells(TrefferLagerliste, 13).Value = auftr�ge.Cells(i, 5).Value 'zu wann
                lagerliste.Cells(TrefferLagerliste, 14).Value = auftr�ge.Cells(i, 6).Value 'wer
                auftr�ge.Rows(i).Delete

            ElseIf Not IsNumeric(auftr�ge.Cells(i, 3).Value) Then
                auftr�ge.Cells(i, 10).Value = "Spalte 'Wieviel' ist keine Zahl"
                Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
            ElseIf auftr�ge.Cells(i, 1).Value = "Bestellt" Then      'neue aktion Bestellt f�rs reportfunktion
                If lagerliste.Cells(TrefferLagerliste, 15).Value + auftr�ge.Cells(i, 3).Value < 0 Then
                    auftr�ge.Cells(i, 10).Value = "Buchung w�rde zu negativer Bestellmenge f�hren"
                    Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                Else
                    Call journalEintragAnlegen(journal, TrefferLagerliste)
                    journal.Cells(1, 4).Value = "Bestellt-Menge ge�ndert"
                    journal.Cells(1, 5).Value = "von " & lagerliste.Cells(TrefferLagerliste, 15).Value
                    lagerliste.Cells(TrefferLagerliste, 15).Value = lagerliste.Cells(TrefferLagerliste, 15).Value + auftr�ge.Cells(i, 3).Value
                    lagerliste.Cells(TrefferLagerliste, 11).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss") 'letzte bewegung
                    journal.Cells(1, 6).Value = "auf " & lagerliste.Cells(TrefferLagerliste, 15).Value
                    lagerliste.Cells(TrefferLagerliste, 14).Value = auftr�ge.Cells(i, 6).Value 'wer in lagerliste eintragen
                    journal.Cells(1, 7).Value = auftr�ge.Cells(i, 6).Value 'wer in journal �bertragen
                    auftr�ge.Rows(i).Delete
                End If
            ElseIf auftr�ge.Cells(i, 2).Value = "Lager" And auftr�ge.Cells(i, 1).Value = "Bedarf" Then      'neue aktion bedarf f�rs lager
                If lagerliste.Cells(TrefferLagerliste, 10).Value + auftr�ge.Cells(i, 3).Value < 0 Then
                    auftr�ge.Cells(i, 10).Value = "Buchung w�rde zu negativem Lagerbedarf f�hren"
                    Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                Else
                    Call journalEintragAnlegen(journal, TrefferLagerliste)
                    journal.Cells(1, 4).Value = "Lagerbedarf ge�ndert"
                    journal.Cells(1, 5).Value = "von " & lagerliste.Cells(TrefferLagerliste, 10).Value
                    lagerliste.Cells(TrefferLagerliste, 10).Value = lagerliste.Cells(TrefferLagerliste, 10).Value + auftr�ge.Cells(i, 3).Value
                    lagerliste.Cells(TrefferLagerliste, 12).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
                    journal.Cells(1, 6).Value = "auf " & lagerliste.Cells(TrefferLagerliste, 10).Value
                    lagerliste.Cells(TrefferLagerliste, 13).Value = auftr�ge.Cells(i, 5).Value 'zu wann
                    lagerliste.Cells(TrefferLagerliste, 14).Value = auftr�ge.Cells(i, 6).Value 'wer in lagerliste eintragen
                    journal.Cells(1, 7).Value = auftr�ge.Cells(i, 6).Value 'wer in journal �bertragen
                    auftr�ge.Rows(i).Delete
                End If
            ElseIf auftr�ge.Cells(i, 2).Value = "Lager" And auftr�ge.Cells(i, 1).Value = "Einkauf" Then
                auftr�ge.Cells(i, 10).Value = "Paarung Einkauf-Lager nicht vorgesehen, bitte Bestand-Lager benutzen"
                Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
            ElseIf auftr�ge.Cells(i, 1).Value = "Bestand" And auftr�ge.Cells(i, 2).Value = "Lager" Then       'Fallunterscheidung Lager (Nachbestellen vs Anzahl mit Plausibilit�tspr�fung)
                If lagerliste.Cells(TrefferLagerliste, 9).Value = "Nachbestellen" Then  'falls erste buchung nach bestellmarkierung
                    Call journalEintragAnlegen(journal, TrefferLagerliste)
                    journal.Cells(1, 4).Value = "Lagerbestand Roter Punkt wieder aufgef�llt"
                    journal.Cells(1, 5).Value = ""
                    lagerliste.Cells(TrefferLagerliste, 9).Value = auftr�ge.Cells(i, 3).Value
                    
                    lagerliste.Cells(TrefferLagerliste, 11).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
                    journal.Cells(1, 6).Value = "auf " & auftr�ge.Cells(i, 3).Value
                    journal.Cells(1, 7).Value = auftr�ge.Cells(i, 6).Value 'wer
                    auftr�ge.Rows(i).Delete
                ElseIf lagerliste.Cells(TrefferLagerliste, 9).Value + auftr�ge.Cells(i, 3).Value < 0 Then
                    auftr�ge.Cells(i, 10).Value = "Buchung w�rde zu negativem Lagerbestand f�hren"
                    Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                Else 'lagerbestand �ndern
                    Call journalEintragAnlegen(journal, TrefferLagerliste)
                    journal.Cells(1, 4).Value = "Lagerbestand ge�ndert"
                    journal.Cells(1, 5).Value = "von " & lagerliste.Cells(TrefferLagerliste, 9).Value
                    lagerliste.Cells(TrefferLagerliste, 9).Value = lagerliste.Cells(TrefferLagerliste, 9).Value + auftr�ge.Cells(i, 3).Value
                    
                    lagerliste.Cells(TrefferLagerliste, 11).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
                    journal.Cells(1, 6).Value = "auf " & lagerliste.Cells(TrefferLagerliste, 9).Value
                    journal.Cells(1, 7).Value = auftr�ge.Cells(i, 6).Value 'wer
                    auftr�ge.Rows(i).Delete
                End If
            ElseIf (auftr�ge.Cells(i, 1).Value = "Bestand") Or (auftr�ge.Cells(i, 1).Value = "Bedarf") Or (auftr�ge.Cells(i, 1).Value = "Einkauf") Then 'projekt <-->bestand/bedarf bewegung
                Dim spalteProjekt As Integer
                Dim strJournal As String
                Dim Lagerbuchung As Boolean
                If auftr�ge.Cells(i, 1).Value = "Bestand" Then 'Projekt <--> Lager bewegungen (Bereitstellung)
                    spalteProjekt = 7
                    strJournal = "Bestand"
                    Lagerbuchung = True
                ElseIf auftr�ge.Cells(i, 1).Value = "Einkauf" Then
                    spalteProjekt = 7
                    strJournal = "Einkauf"
                    Lagerbuchung = False
                Else                                           'Bedarfsver�nderung
                    spalteProjekt = 8
                    strJournal = "Bedarf"
                    Lagerbuchung = False
                End If
                
                If Workbooks(b).Sheets.Count > 1 Then 'Pr�fen ob Ziel als tabellenblattname gefunden wird
                    
                    Dim Blattgefunden As Boolean
                    Blattgefunden = False
                    For k = 2 To Workbooks(b).Sheets.Count
                        Dim Projektliste As Worksheet
                        Set Projektliste = Workbooks(b).Worksheets(k)
                        If Projektliste.Name = auftr�ge.Cells(i, 2).Value Then  'falls ja suchen ob EAN vorhanden,
                            Blattgefunden = True
                            Dim checkLager As Integer
                            If Lagerbuchung = True Then
                                checkLager = lagerliste.Cells(TrefferLagerliste, 9).Value - auftr�ge.Cells(i, 3).Value
                            Else
                                checkLager = 1
                            End If
                            Set treffer = Nothing
                            Set treffer = Projektliste.Columns(1).Find(what:=auftr�ge.Cells(i, 4).Value, lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
                            If treffer Is Nothing Then 'falls nicht vorhanden, bestnd/bedarf neu anlegen
                                If auftr�ge.Cells(i, 3).Value < 0 Then 'wenn wiviel < 0 fehler (abgang vom projekt)
                                    auftr�ge.Cells(i, 10).Value = "Dies w�rde zu negativem Projekt-" & strJournal & " f�hren"
                                    Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                                ElseIf checkLager < 0 Then 'auf plausibilit�t bei abgang vom lager pr�fen
                                    auftr�ge.Cells(i, 10).Value = "Buchung w�rde zu negativem Lagerbestand f�hren"
                                    Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                                Else 'dann buchen in neuer projektzeile
                                    'oben einf�gen und neu anlegen
                                    Projektliste.Rows(2).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
                                    Call journalEintragAnlegen(journal, TrefferLagerliste)
                                    journal.Cells(1, 4).Value = "Projekt-" & strJournal & " " & auftr�ge.Cells(i, 2).Value & " ertsmals angelegt"
                                    journal.Cells(1, 5).Value = "von 0"
                                    'Lagerliste �ndern
                                    If Lagerbuchung = True Then
                                        lagerliste.Cells(TrefferLagerliste, 9).Value = lagerliste.Cells(TrefferLagerliste, 9).Value - auftr�ge.Cells(i, 3).Value
                                        
                                        lagerliste.Cells(TrefferLagerliste, 11).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
                                    End If
                                    'Teileeigenschaften aus lagerliste kopieren
                                    lagerliste.Range(lagerliste.Cells(TrefferLagerliste, 1), lagerliste.Cells(TrefferLagerliste, 6)).Copy Destination:=Projektliste.Cells(2, 1) '1-6 kopieren
                                    'Projektbestand/bedarf neu anlegen und mit Zeitstempel versehen
                                    Projektliste.Cells(2, 11).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
                                    Projektliste.Cells(2, spalteProjekt).Value = auftr�ge.Cells(i, 3).Value
                                    'falls bedarfsbuchung, dann wann spalte und zus�tzichen zeitstempel
                                    If strJournal = "Bedarf" Then
                                        Projektliste.Cells(2, 10).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")   'letzte bedarfsmeldung
                                        Projektliste.Cells(2, 9).Value = auftr�ge.Cells(i, 5).Value 'zu wann
                                        Projektliste.Cells(2, 12).Value = auftr�ge.Cells(i, 6).Value 'wer in projektliste eintragen
                                    End If
                                    'ins journal �bernehmen
                                    journal.Cells(1, 6).Value = "auf " & Projektliste.Cells(2, spalteProjekt).Value
                                    journal.Cells(1, 7).Value = auftr�ge.Cells(i, 6).Value 'wer
                                    auftr�ge.Rows(i).Delete
                                End If
                            Else 'scancode vorhanden
                                Dim trefferProjektzeile As Integer
                                trefferProjektzeile = treffer.Row
                                    
                                If Projektliste.Cells(trefferProjektzeile, spalteProjekt).Value + auftr�ge.Cells(i, 3).Value < 0 Then 'auf plausibilit�t bei abgang vom projekt pr�fen
                                    auftr�ge.Cells(i, 10).Value = "Dies w�rde zu negativem Projekt-" & strJournal & " f�hren"
                                    Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                                ElseIf checkLager < 0 Then 'auf plausibilit�t bei abgang vom Lager pr�fen
                                    auftr�ge.Cells(i, 10).Value = "Buchung w�rde zu negativem Lagerbestand f�hren"
                                    Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                                Else 'dann buchen in gefundener projektzeile
                                    Call journalEintragAnlegen(journal, TrefferLagerliste)
                                    journal.Cells(1, 4).Value = "Projekt-" & strJournal & " " & auftr�ge.Cells(i, 2).Value & " ge�ndert"
                                    journal.Cells(1, 5).Value = "von " & Projektliste.Cells(trefferProjektzeile, spalteProjekt).Value
                                    'Lagerliste �ndern
                                    If Lagerbuchung = True Then
                                        lagerliste.Cells(TrefferLagerliste, 9).Value = lagerliste.Cells(TrefferLagerliste, 9).Value - auftr�ge.Cells(i, 3).Value
                                        
                                        lagerliste.Cells(TrefferLagerliste, 11).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
                                    End If
                                    'Projektbestand/bedarf �ndern und mit Zeitstempel versehen
                                    Projektliste.Cells(trefferProjektzeile, spalteProjekt).Value = Projektliste.Cells(trefferProjektzeile, spalteProjekt).Value + auftr�ge.Cells(i, 3).Value
                                    Projektliste.Cells(trefferProjektzeile, 11).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
                                        'falls bedarfsbuchung, dann wann spalte und zus�tzichen zeitstempel
                                   If strJournal = "Bedarf" Then
                                        Projektliste.Cells(trefferProjektzeile, 10).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
                                        Projektliste.Cells(trefferProjektzeile, 9).Value = auftr�ge.Cells(i, 5).Value
                                        Projektliste.Cells(trefferProjektzeile, 12).Value = auftr�ge.Cells(i, 6).Value 'wer in projektliste eintragen
                                    End If
                                    'ins journal �bernehmen
                                    journal.Cells(1, 6).Value = "auf " & Projektliste.Cells(trefferProjektzeile, spalteProjekt).Value
                                    journal.Cells(1, 7).Value = auftr�ge.Cells(i, 6).Value 'wer
                                    auftr�ge.Rows(i).Delete
                                End If
                            End If
                            Exit For 'abbruch der schleife nach erster gefundenen Projektmappe
                        End If
                    Next k
                    If Not Blattgefunden Then
                        auftr�ge.Cells(i, 10).Value = "Projektnummer wurde nicht als Tabellenblatt gefunden"
                        Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                    End If
                Else
                    auftr�ge.Cells(i, 10).Value = "keine Tabellenbl�tter mit Projektbestand vorhanden"
                    Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
                End If
            Else
               auftr�ge.Cells(i, 10).Value = "Unbehandelter Buchungstyp"
               Call Problemkinder(auftr�ge, journal, i) 'ausschneiden zu problemkinder
            End If
n�chsteZeile:
        Next i
        'On Error GoTo FehlerBeimSpeichern
        Workbooks(b).Close SaveChanges:=True
        Workbooks(c).Close SaveChanges:=True
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
         '   Beep 2000, 750
        End If
        Workbooks(Dateiname).Save
    End If
   
    
Exit Sub


LagerlisteSchonOffen:
    Workbooks(b).Close SaveChanges:=False
    MsgBox "Zur Zeit nicht m�glich, Lagerliste wird gerade verwendet"
Exit Sub

JournalSchonOffen:
    MsgBox "Zur Zeit nicht m�glich, Journal wird gerade verwendet"
    Workbooks(c).Close SaveChanges:=False
Exit Sub
FehlerBeimSpeichern:
    MsgBox "Fehler beim finalen Speichern. Bitte folgende Codes notieren, alle Fenster bis auf das Terminal schlie�en, dann Buchungsliste neu anlegen und TK informieren!" & vbCrLf & Err.Number & vbCrLf & Err.Description
Exit Sub
FehlerOpenLagerliste:
    MsgBox "Fehler beim �ffnen Lagerliste. Bitte folgende Codes notieren und TK informieren!" & vbCrLf & Err.Number & vbCrLf & Err.Description
Exit Sub
FehlerOpenJournal:
    MsgBox "Fehler beim �ffnen Journal. Bitte folgende Codes notieren und TK informieren!" & vbCrLf & Err.Number & vbCrLf & Err.Description
    Workbooks(b).Close SaveChanges:=False
Exit Sub

FehlerVorBuchung:
    MsgBox "!!!KRITISCH!!! Fehler beim Liste abarbeiten vor dem Speichern, es k�nnen Buchungen aus der Liste verloren gegangen sein. Bitte folgende Codes notieren, alle Fenster bis auf das Terminal schlie�en, dann Buchungsliste neu anlegen und TK informieren!" & vbCrLf & Err.Number & vbCrLf & Err.Description
    On Error Resume Next
    Workbooks(b).Close SaveChanges:=False
    Workbooks(c).Close SaveChanges:=False

End Sub

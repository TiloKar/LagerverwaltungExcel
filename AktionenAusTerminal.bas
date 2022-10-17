Attribute VB_Name = "AktionenAusTerminal"
Public Function checkBedarf(aktion, projektnummer, Scancode)
    checkBedarf = "nein"
    For i = 0 To plcount - 1                                                                                                        'Vom ersten bis zum letzten Projekt durchgehen
        If plnames(i) = projektnummer Then
            For k = 2 To plrows(i)
               If StrComp(Scancode, pl(i)(k, 1), 1) = 0 Then                                                                        'Von Projektnummer aus Projektliste mit Scancode übereinstimmt
                  If IsNumeric(pl(i)(k, 8)) And aktion = "Bedarf" Then
                    If pl(i)(k, 8) > 0 Then
                        checkBedarf = "Bedarf von " & pl(i)(k, 8) & " bereits eingetragen mit Datum vom " & pl(i)(k, 10) & "."
                    End If
                  End If
                  If IsNumeric(pl(i)(k, 7)) Then
                    If pl(i)(k, 7) > 0 Then
                        If checkBedarf = "nein" Then
                            checkBedarf = "Bestand von " & pl(i)(k, 7) & " bereits in der Bereitstellung " & "."
                        Else
                            checkBedarf = checkBedarf & " Bestand von " & pl(i)(k, 7) & " bereits in der Bereitstellung " & "."
                        End If
                    End If
                  End If
                  Exit For
               End If
            Next k
            Exit For
        End If
    Next i
    
End Function
Public Sub EinträgeLöschen()
    Buchungsterminal.keinTreffer.Caption = ""
    Buchungsterminal.Projektnamen.Caption = ""
    Buchungsterminal.ProjektBedarf.Caption = ""
    Buchungsterminal.ProjektBestellt.Caption = ""
    Buchungsterminal.ProjektBereitstellung.Caption = ""
    Buchungsterminal.nBedarf.Caption = ""
    Buchungsterminal.nBestellt.Caption = ""
    Buchungsterminal.nBereitstellung.Caption = ""
    Buchungsterminal.nLager.Caption = ""
    Buchungsterminal.Bestandsüberwachung.Caption = ""
    Buchungsterminal.Zulieferer.Caption = ""
    Buchungsterminal.Lagerort.Caption = ""
    Buchungsterminal.Bezeichner1.Caption = ""
    Buchungsterminal.Bezeichner2.Caption = ""
    Buchungsterminal.Circuit.Caption = ""
    Buchungsterminal.Value.Caption = ""
    Buchungsterminal.Nachbestellen.Visible = False
End Sub
Public Sub Scancodeauswerten()

    If Buchungsterminal.Scancode.Value = "" Then                                                       'Nix machen falls nix eingetragen wurde
        If Buchungsterminal.keinTreffer.Caption <> "Kein Treffer :-(" And Buchungsterminal.keinTreffer.Caption <> "" Then
            Buchungsterminal.Scancode.Value = Buchungsterminal.keinTreffer.Caption
        Else
            Exit Sub
        End If
    End If
    
    Call EinträgeLöschen
    Dim pstr As String
    'SCANCODES FÜR KNÖPFE
    If InStr(Buchungsterminal.Scancode.Value, "§§1111") = 1 Then                                        'Wenn das der Fall ist dann wird der Einbuchen Knopf getoggelt
        Buchungsterminal.ToggleEinbuchen.Value = True
        Buchungsterminal.Scancode.Value = ""
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then               'initialiserung Beep
         '   Beep 2000, 750
        End If
    ElseIf InStr(Buchungsterminal.Scancode.Value, "§§2222") = 1 Then
        Buchungsterminal.ToggleAusbuchen.Value = True
        Buchungsterminal.Scancode.Value = ""
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
          '  Beep 2000, 750
        End If
    ElseIf InStr(Buchungsterminal.Scancode.Value, "§§3333") = 1 Then
        Buchungsterminal.ToggleSuchen.Value = True
        Buchungsterminal.Scancode.Value = ""
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
          '  Beep 2000, 750
        End If
    ElseIf InStr(Buchungsterminal.Scancode.Value, "§§4444") = 1 Then
        Buchungsterminal.ToggleBestand.Value = True
        Buchungsterminal.Scancode.Value = ""
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
        '    Beep 2000, 750
        End If
    ElseIf InStr(Buchungsterminal.Scancode.Value, "§§5555") = 1 Then
        Buchungsterminal.ToggleBedarf.Value = True
        Buchungsterminal.Scancode.Value = ""
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
        '    Beep 2000, 750
        End If
    ElseIf InStr(Buchungsterminal.Scancode.Value, "##1400") = 1 Then
        Buchungsterminal.Projektauswahl.ListIndex = 0
        Buchungsterminal.Scancode.Value = ""
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
         '   Beep 2000, 750
        End If
    ElseIf InStr(Buchungsterminal.Scancode.Value, "%%") = 1 Then                                            'user wechseln / Scancodes für User
        pstr = Replace(Buchungsterminal.Scancode.Value, "%%", "")                                           'pstr enthält dann den Scancode ersetzt %% mit ''
        For i = 0 To Buchungsterminal.Nutzer.ListCount - 1
            If Buchungsterminal.Nutzer.List(i) = pstr Then
                Buchungsterminal.Nutzer.ListIndex = i
                If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
                 '   Beep 2000, 750
                End If
            End If
        Next i
        
        
    ElseIf InStr(Buchungsterminal.Scancode.Value, "##") = 1 Then
        'Projektnummern durchsuchen
        pstr = Replace(Buchungsterminal.Scancode.Value, "##", "")
        Dim k As Integer
        Dim Index As Integer
        Index = 0
        If Buchungsterminal.Projektauswahl.ListCount > 1 Then
            For k = 1 To Buchungsterminal.Projektauswahl.ListCount - 1
                If Buchungsterminal.Projektauswahl.List(k) = pstr Then
                    Index = k
                    Exit For
                End If
            Next k
        End If
        If Index > 0 Then
            Buchungsterminal.Projektauswahl.ListIndex = Index
            If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
              '  Beep 2000, 750
            End If
        Else
            MsgBox "Projektcode in Datenbank nicht gefunden"
        End If
        
        Buchungsterminal.Scancode.Value = ""
    Else
        
        
        Dim treffer As Integer
        treffer = 0
        For i = 1 To llrows
            If StrComp(Buchungsterminal.Scancode.Value, ll(i, 1), 1) = 0 Then           'Tabellenblatt 1, Spalte 1 Zeile wird durchgegenegn bis zur letzten
                 treffer = i                                                            'Wenn gleich kriegt treffer den Wert der Zeile
                 Exit For
            End If
        Next i
       
        If treffer <= 1 Then                                                             'Falls Header oder gar kein Treffer --> Fehlermeldung
            Buchungsterminal.keinTreffer.Caption = "Kein Treffer :-("
            Buchungsterminal.keinTreffer.ForeColor = &HFF&
        Else
            Call TrefferAnzeigen(treffer)
            If Inventurmodus = True Then
                If Buchungsterminal.ToggleSuchen.Value = False Then
                    Call BuchungInListeAnlegen("Einkauf", Buchungsterminal.Projektauswahl.Text, Buchungsterminal.nBuchungen.Value, Buchungsterminal.keinTreffer.Caption, Buchungsterminal.Wann.Value, Buchungsterminal.Bezeichner1.Caption, Buchungsterminal.Bezeichner2.Caption, Buchungsterminal.Zulieferer.Caption, Buchungsterminal.Nutzer.Text)
                End If
            ElseIf Not IsNumeric(ll(treffer, 9)) And ll(treffer, 9) <> "Nachbestellen" Then
                 MsgBox "Lagerbestand lässt sich nicht als Zahl interpretieren"
            Else
                 If scan = False And userlevel = 0 Then  'lagerterminal roter punkt
                    'nix
                 ElseIf scan = False And Buchungsterminal.ToggleSuchen.Value = False Then                   'mädels roter punkt
                     Call BuchungInListeAnlegen("Bestand", "Lager", Buchungsterminal.nBuchungen.Value, Buchungsterminal.keinTreffer.Caption, Buchungsterminal.Wann.Value, Buchungsterminal.Bezeichner1.Caption, Buchungsterminal.Bezeichner2.Caption, Buchungsterminal.Zulieferer.Caption, Buchungsterminal.Nutzer.Text)
                     Call BuchungInListeAnlegen("Bestellt", "x", -Buchungsterminal.nBuchungen.Value, Buchungsterminal.keinTreffer.Caption, "", Buchungsterminal.Bezeichner1.Caption, Buchungsterminal.Bezeichner2.Caption, Buchungsterminal.Zulieferer.Caption, Buchungsterminal.Nutzer.Text) 'zusätzlich bestellt menge mit umgekehrtem vorzeichen
                 ElseIf Buchungsterminal.ToggleSuchen.Value = False Then   'nur scanbare artikel können gebucht werden
                     If Buchungsterminal.ToggleBestand.Value = True Then
                        
                        If Buchungsterminal.Einkaufsbuchung.Value = True Then
                            Call BuchungInListeAnlegen("Einkauf", Buchungsterminal.Projektauswahl.Text, Buchungsterminal.nBuchungen.Value, Buchungsterminal.keinTreffer.Caption, Buchungsterminal.Wann.Value, Buchungsterminal.Bezeichner1.Caption, Buchungsterminal.Bezeichner2.Caption, Buchungsterminal.Zulieferer.Caption, Buchungsterminal.Nutzer.Text)
                            Call BuchungInListeAnlegen("Bestellt", "x", -Buchungsterminal.nBuchungen.Value, Buchungsterminal.keinTreffer.Caption, "", Buchungsterminal.Bezeichner1.Caption, Buchungsterminal.Bezeichner2.Caption, Buchungsterminal.Zulieferer.Caption, Buchungsterminal.Nutzer.Text) 'zusätzlich bestellt menge mit umgekehrtem vorzeichen
                        Else
                            Call BuchungInListeAnlegen("Bestand", Buchungsterminal.Projektauswahl.Text, Buchungsterminal.nBuchungen.Value, Buchungsterminal.keinTreffer.Caption, Buchungsterminal.Wann.Value, Buchungsterminal.Bezeichner1.Caption, Buchungsterminal.Bezeichner2.Caption, Buchungsterminal.Zulieferer.Caption, Buchungsterminal.Nutzer.Text)
                            If Buchungsterminal.Projektauswahl.Text = "Lager" And userlevel > 0 Then
                                Call BuchungInListeAnlegen("Bestellt", "x", -Buchungsterminal.nBuchungen.Value, Buchungsterminal.keinTreffer.Caption, "", Buchungsterminal.Bezeichner1.Caption, Buchungsterminal.Bezeichner2.Caption, Buchungsterminal.Zulieferer.Caption, Buchungsterminal.Nutzer.Text) 'zusätzlich bestellt menge mit umgekehrtem vorzeichen
                            End If
                        End If
                     Else
                         Call BuchungInListeAnlegen("Bedarf", Buchungsterminal.Projektauswahl.Text, Buchungsterminal.nBuchungen.Value, Buchungsterminal.keinTreffer.Caption, Buchungsterminal.Wann.Value, Buchungsterminal.Bezeichner1.Caption, Buchungsterminal.Bezeichner2.Caption, Buchungsterminal.Zulieferer.Caption, Buchungsterminal.Nutzer.Text)
                     End If
                 End If
            End If
           Buchungsterminal.Scancode.Value = ""
           Buchungsterminal.Scancode.SetFocus
        End If
    End If
    'MsgBox ("hier 2")
End Sub

Public Sub TrefferAnzeigen(Trefferzeile As Integer)
    
    Buchungsterminal.keinTreffer.Caption = ll(Trefferzeile, 1)
    Buchungsterminal.nLager.Caption = ll(Trefferzeile, 9)
    Buchungsterminal.keinTreffer.ForeColor = &HC000&
    Buchungsterminal.Bezeichner1.Caption = ll(Trefferzeile, 2)          'Bezeichner 1
    Buchungsterminal.Bezeichner2.Caption = ll(Trefferzeile, 3)          'Bezeichner 1
    Buchungsterminal.Zulieferer.Caption = ll(Trefferzeile, 4)           'Zulieferer
    Buchungsterminal.Circuit.Caption = ll(Trefferzeile, 5)              'Circuit
    Buchungsterminal.Value.Caption = ll(Trefferzeile, 6)                'Value
    Buchungsterminal.Lagerort.Caption = ll(Trefferzeile, 7)             'Lagerort
    Buchungsterminal.nBestellt.Caption = ll(Trefferzeile, 15)           'Bestellt Anzahl

    If (ll(Trefferzeile, 8) = "Nein" Or ll(Trefferzeile, 8) = "NEIN" Or ll(Trefferzeile, 8) = "nein") And (Inventurmodus = False) Then
        'Steuerelemente für Roter Punkt verstecken
        scan = False
        Buchungsterminal.ToggleAusbuchen.Visible = False
        Buchungsterminal.ToggleBestand.Visible = False
        Buchungsterminal.ToggleBedarf.Visible = False
        Buchungsterminal.Projektauswahl.Visible = False
        Buchungsterminal.HintProjektauswahl1.Visible = False
        Buchungsterminal.HintProjektauswahl2.Visible = False
        Buchungsterminal.Wann.Visible = True
        Buchungsterminal.HLWann.Visible = True
        Buchungsterminal.HintWann.Visible = True
        
        'einbuchen nur für mädels
        If user = "Lager" Then
            Buchungsterminal.ToggleEinbuchen.Visible = False             'Verstecken der Knöpfe - 'Userlvl'
            Buchungsterminal.HLProjektauswahl.Visible = False
        Else
            Buchungsterminal.ToggleEinbuchen.Visible = True
            Buchungsterminal.Projektauswahl.Visible = True
            Buchungsterminal.Projektauswahl.ListIndex = 0
        End If
        'Hinweis auf Scanart
        Buchungsterminal.Bestandsüberwachung.ForeColor = &HFF&
        If CStr(ll(Trefferzeile, 9)) <> "Nachbestellen" Then
            Buchungsterminal.Nachbestellen.Visible = True
            Buchungsterminal.Bestandsüberwachung.Caption = "Scannen nicht nötig!"
        Else
            
            Buchungsterminal.Nachbestellen.Visible = False
            Buchungsterminal.Bestandsüberwachung.Caption = "Scannen nicht nötig! Best. vorgemerkt"
        End If
    Else
        Buchungsterminal.Bestandsüberwachung.Caption = "genaues Scannen nötig!"
        Buchungsterminal.Bestandsüberwachung.ForeColor = &HC000&
        Buchungsterminal.Nachbestellen.Visible = False
        Buchungsterminal.ToggleAusbuchen.Visible = True
        Buchungsterminal.ToggleBestand.Visible = True
        Buchungsterminal.ToggleBedarf.Visible = True
        Buchungsterminal.Projektauswahl.Visible = True
        Buchungsterminal.ToggleEinbuchen.Visible = True
         Buchungsterminal.HintProjektauswahl1.Visible = True
        Buchungsterminal.HintProjektauswahl2.Visible = True
        Buchungsterminal.HLProjektauswahl.Visible = True
        scan = True
    End If
    
    
     ' suchen nach treffer in projekten
     Dim i As Integer
     Dim strname As String
     Dim strbedarf As String
     Dim strbereit As String
     Dim strbestellt As String
     Dim SumBedarf As Long
     Dim SumBereit As Long
     strname = ""
     strbedarf = ""
     strbereit = ""
     strbestellt = ""
     SumBedarf = 0
     SumBereit = 0

    'neu auch lagerliste nach bedarf durchsuchen
    Dim treffer As Integer
    treffer = 0
    For k = 2 To llrows
        If StrComp(Buchungsterminal.keinTreffer.Caption, ll(k, 1), 1) = 0 Then
             treffer = k
             Exit For
        End If
    Next k
    If treffer > 1 Then
        strname = strname & "Lager" & vbCrLf
        If IsNumeric(ll(treffer, 10)) Then
            SumBedarf = SumBedarf + ll(treffer, 10)
            strbedarf = strbedarf & CStr(ll(treffer, 10)) & vbCrLf              'vbCrLf = Zeilenumbruch
        Else
            MsgBox "Bedarf in Lagerliste lässt sich nicht als Zahl interpretieren"
            strbedarf = strbedarf & "?" & vbCrLf
        End If
         If IsNumeric(ll(treffer, 9)) Then
            SumBereit = SumBereit + ll(treffer, 9)
            strbereit = strbereit & CStr(ll(treffer, 9)) & vbCrLf
        Else
            strbereit = strbereit & "?" & vbCrLf
        End If
        If IsNumeric(ll(treffer, 9)) And IsNumeric(ll(treffer, 10)) Then
            If ll(treffer, 10) - ll(treffer, 9) > 0 Then
                strbestellt = strbestellt & "offen" & vbCrLf
            Else
                strbestellt = strbestellt & ":-)" & vbCrLf
            End If
        Else
           strbestellt = strbestellt & "?" & vbCrLf
        End If
        
        
    End If
    If plcount > 0 Then
         For i = 0 To plcount - 1
             treffer = 0
             For k = 2 To plrows(i)
                If StrComp(Buchungsterminal.keinTreffer.Caption, pl(i)(k, 1), 1) = 0 Then               'Wenn keinTreffer.Caption = der Projektliste i in Zeile k Spalte 1 ist dann ist treffer=Zeile
                     treffer = k
                     Exit For
                End If
             Next k

             If treffer > 1 Then
                 strname = strname & Buchungsterminal.Projektauswahl.List(i + 1) & vbCrLf
                 If IsNumeric(pl(i)(treffer, 8)) Then
                     SumBedarf = SumBedarf + pl(i)(treffer, 8)
                     strbedarf = strbedarf & CStr(pl(i)(treffer, 8)) & vbCrLf
                 Else
                     MsgBox "Bedarf in Projekt " & Buchungsterminal.Projektauswahl.List(i) & " lässt sich nicht als Zahl interpretieren"
                     strbedarf = strbedarf & "?" & vbCrLf
                 End If
                 If IsNumeric(pl(i)(treffer, 7)) Then
                     SumBereit = SumBereit + pl(i)(treffer, 7)
                     strbereit = strbereit & CStr(pl(i)(treffer, 7)) & vbCrLf
                 Else
                     MsgBox "Bestand in Projekt " & Buchungsterminal.Projektauswahl.List(i) & " lässt sich nicht als Zahl interpretieren"
                     strbereit = strbereit & "?" & vbCrLf
                 End If
                 If IsNumeric(pl(i)(treffer, 7)) And IsNumeric(pl(i)(treffer, 8)) Then
                    If pl(i)(treffer, 8) - pl(i)(treffer, 7) > 0 Then
                        strbestellt = strbestellt & "offen" & vbCrLf
                    Else
                        strbestellt = strbestellt & ":-)" & vbCrLf
                    End If
                Else
                   strbestellt = strbestellt & "?" & vbCrLf
                End If
                 
             End If
         Next i
         
         
         Buchungsterminal.Projektnamen.Caption = strname
         Buchungsterminal.ProjektBereitstellung.Caption = strbereit
         Buchungsterminal.ProjektBedarf.Caption = strbedarf
         Buchungsterminal.nBedarf.Caption = SumBedarf
         Buchungsterminal.nBereitstellung.Caption = SumBereit
         Buchungsterminal.ProjektBestellt.Caption = strbestellt
         
     End If
End Sub
Public Sub NeueBuchung(terminal, aktion, ziel, wieviel, ean, Wann, bez1, bez2, zul, wer)
    terminal.Rows(2).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    terminal.Rows(2).Font.Color = RGB(0, 0, 0)
    terminal.Cells(2, 1).Value = aktion
    terminal.Cells(2, 2).Value = ziel
    If CStr(wieviel) = "Nachbestellen" Then
        terminal.Cells(2, 3).Value = "Nachbestellen"
    ElseIf CStr(wieviel) = "?" Then
        terminal.Cells(2, 3).Value = "?"
    ElseIf Buchungsterminal.ToggleEinbuchen.Value = True Then   ' Suchen triggert gar nicht erst dieses prozedur
        terminal.Cells(2, 3).Value = Application.WorksheetFunction.Round(CSng(wieviel), nachkomma)
    Else
        terminal.Cells(2, 3).Value = Application.WorksheetFunction.Round(CSng(wieviel) * -1, nachkomma)
    End If
    terminal.Cells(2, 4).Value = ean 'Buchungsterminal.keinTreffer.Caption
    terminal.Cells(2, 5).Value = Wann 'Buchungsterminal.wann.Value
    terminal.Cells(2, 7).Value = bez1 'Buchungsterminal.Bezeichner1.Caption
    terminal.Cells(2, 8).Value = bez2 'Buchungsterminal.Bezeichner2.Caption
    terminal.Cells(2, 9).Value = zul 'Buchungsterminal.Zulieferer.Caption
    terminal.Cells(2, 6).Value = wer
    If (aktion = "Bedarf" Or aktion = "Bestand") And ziel <> "Lager" Then
       bedarfschonvorhanden = checkBedarf(aktion, ziel, ean)
       If bedarfschonvorhanden <> "nein" Then
            terminal.Cells(2, 10).Value = bedarfschonvorhanden
       End If
    End If
    
End Sub

Public Function FindeTrefferInBuchungsliste(terminal, aktion, ziel, wieviel, ean)
    FindeTrefferInBuchungsliste = 0
    Dim i As Integer
    Dim treffer As Object
    Set treffer = Nothing
    Set treffer = terminal.Columns(1).Find(what:="", lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
    If Not (treffer Is Nothing) Then
        If treffer.Row > 2 Then
            For i = 2 To (treffer.Row - 1)
                Dim sp1 As Boolean
                Dim sp2 As Boolean
                Dim sp3 As Boolean
                Dim sp4 As Boolean
                sp1 = CStr(terminal.Cells(i, 1).Value) = aktion
                sp2 = CStr(terminal.Cells(i, 2).Value) = ziel      'wichtig damit projektnummern auch als string verglichen werden
                sp3 = StrComp(terminal.Cells(i, 4).Value, ean, 1) = 0
                If Not IsNumeric(terminal.Cells(i, 3).Value) And Not IsNumeric(wieviel) Then    ' auf nachbesttellen prüfen
                    sp4 = CStr(terminal.Cells(i, 3).Value) = CStr(wieviel)
                Else
                    sp4 = True          'sonst wiviel ignorieren
                End If
                'MsgBox sp1 & sp2 & sp3 & sp4
                'MsgBox i & Terminal.Cells(i, 3).Value & wieviel
                If sp1 And sp2 And sp3 And sp4 Then
                    FindeTrefferInBuchungsliste = i
                    Exit For
                End If
            Next i
        End If
    End If
    
End Function

Public Sub BuchungInListeAnlegen(aktion, ziel, wieviel, ean, Wann, bez1, bez2, zul, wer)
    
    If EsWurdeAufButtonGescannt Then
        'nix            to do überprüfen
        EsWurdeAufButtonGescannt = False
    ElseIf wer = "" Then
        MsgBox ("Bitte Nutzer wählen")
    ElseIf ziel = "Lager" And userlevel = 0 And CStr(wieviel) <> "Nachbestellen" Then
        MsgBox ("Nutzer " & user & " ist nicht für Lagerbuchungen freigegeben")
    ElseIf Not IsNumeric(wieviel) And wieviel <> "Nachbestellen" And wieviel <> "?" Then
        MsgBox "keine gültige Buchungsanzahl eingegegeben"
    ElseIf wieviel = 0 Then
     'nix
    Else
        Application.CutCopyMode = False
        Dim terminal As Worksheet        'temporäre arbeitsblatt
        Set terminal = Workbooks(Dateiname).Worksheets(1)
        Dim treffer As Integer
        treffer = FindeTrefferInBuchungsliste(terminal, aktion, ziel, wieviel, ean)
        If treffer = 0 Then
            Call NeueBuchung(terminal, aktion, ziel, wieviel, ean, Wann, bez1, bez2, zul, wer)
        ElseIf CStr(wieviel) = "Nachbestellen" Then
            terminal.Cells(treffer, 3).Value = "Nachbestellen"
        ElseIf CStr(wieviel) = "?" Then
            terminal.Cells(treffer, 3).Value = "?"
        ElseIf IsNumeric(terminal.Cells(treffer, 3).Value) Then
            If Wann <> "" Then
                terminal.Cells(treffer, 5).Value = Wann
            End If
            If (aktion = "Bedarf" Or aktion = "Bestand") And ziel <> "Lager" Then
               bedarfschonvorhanden = checkBedarf(aktion, ziel, ean)
               If bedarfschonvorhanden <> "nein" Then
                    terminal.Cells(treffer, 10).Value = bedarfschonvorhanden
               End If
            End If
            If Buchungsterminal.ToggleEinbuchen.Value = True Then   ' Suchen triggert gar nicht erst dieses prozedur
                terminal.Cells(treffer, 3).Value = CSng(terminal.Cells(treffer, 3).Value) + CSng(wieviel)
            Else
                terminal.Cells(treffer, 3).Value = CSng(terminal.Cells(treffer, 3).Value) - CSng(wieviel)
            End If
            terminal.Cells(treffer, 6).Value = CStr(wer) 'letzten nutzer eintragen
            terminal.Cells(treffer, 3).Value = Application.WorksheetFunction.Round(terminal.Cells(treffer, 3).Value, nachkomma)
        Else
           MsgBox "'Wieviel' im bestehender Buchungszeile lässt sich nicht interpretieren"
        End If
        
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" And Not keinPiep Then
           ' Beep 2000, 750
        End If
      
    End If

End Sub

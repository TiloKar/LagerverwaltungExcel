Attribute VB_Name = "EigenschaftenBearbeiten"
Public Sub EigenschaftSpeichern()
    Workbooks.Open Filename:=a & b, ReadOnly:=False, WriteResPassword:=pwlager, Password:=pwlager
    If Workbooks(b).ReadOnly Then
        GoTo LagerlisteSchonOffen
    End If
    Dim lagerliste As Worksheet 'temporäre arbeitsblatt
    Set lagerliste = Workbooks(b).Worksheets(1)
    
    Workbooks.Open Filename:=a & c, ReadOnly:=False, WriteResPassword:=pwjournal, Password:=pwjournal
    If Workbooks(c).ReadOnly Then
        GoTo JournalSchonOffen
    End If
    On Error GoTo FehlerNachOpenJournal
    Dim journal As Worksheet        'temporäre arbeitsblatt
    Set journal = Workbooks(c).Worksheets(1)
    'EAN in Lagergerliste finden
    Dim treffer As Object
    Set treffer = Nothing
    Set treffer = lagerliste.Columns(1).Find(what:=Buchungsterminal.keinTreffer.Caption, lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
    If Buchungsterminal.keinTreffer.Caption = "" Or treffer Is Nothing Then
        MsgBox "kein gültiger Barcode ausgewählt"
        Workbooks(b).Close SaveChanges:=False
        Workbooks(c).Close SaveChanges:=False
    Else
        Dim TrefferLagerliste As Integer
        TrefferLagerliste = treffer.Row
        journal.Rows(1).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
 
        journal.Cells(1, 1).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss") 'Format(Now, "YYYY" & "_" & "MM" & "_" & "DD" & "_" & "hh" & "_" & "mm" & "_" & "ss")
        journal.Cells(1, 2).Value = Buchungsterminal.keinTreffer.Caption
        journal.Cells(1, 3).Value = "Eigenschaft"
        journal.Cells(1, 4).Value = Edit.HLEdit.Caption & " geändert"
        journal.Cells(1, 5).Value = "von " & lagerliste.Cells(TrefferLagerliste, Sucheigenschaft).Value
        
        lagerliste.Cells(TrefferLagerliste, Sucheigenschaft).Value = Edit.NeuerText.Value
        
        journal.Cells(1, 6).Value = "auf " & lagerliste.Cells(TrefferLagerliste, Sucheigenschaft).Value
        
        Buchungsterminal.Scancode.Value = Buchungsterminal.keinTreffer.Caption
        
        Workbooks(b).Close SaveChanges:=True
        Workbooks(c).Close SaveChanges:=True
        EigenschaftGeändert = True
        'Call Scancodeauswerten  workaround über mouseover event und triggervariable um  einmal zu aktualisieren
        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
         '   Beep 2000, 750
        End If
        Edit.Hide
        
    End If

Exit Sub
FehlerNachOpenJournal:
    MsgBox "Fehler beim Ändern einer Eigenschaft. Bitte folgende Codes notieren und TK informieren!" & vbCrLf & Err.Number & vbCrLf & Err.Description
    Workbooks(b).Close SaveChanges:=True
    Workbooks(c).Close SaveChanges:=True

JournalSchonOffen:
    MsgBox "Zur Zeit nicht möglich, Journal wird gerade verwendet"
    Workbooks(c).Close SaveChanges:=False
Exit Sub
LagerlisteSchonOffen:
    Workbooks(b).Close SaveChanges:=False
    MsgBox "Zur Zeit nicht möglich, Lagerliste wird gerade verwendet"

End Sub

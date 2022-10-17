Attribute VB_Name = "VorbereitungenForms"
Public Sub Inits()
    Dateiname = ActiveWorkbook.Name
    
   'If Workbooks(Dateiname).Worksheets(3).Cells(8, 4).Value <> "" Then               'öffnen beim start?
    '   a = ThisWorkbook.Path + "\SERVER\" 'workaround für offline - variante
    '    Workbooks(Dateiname).Worksheets(3).Cells(7, 6).Value = a
    'Else
        a = Workbooks(Dateiname).Worksheets(3).Cells(7, 6).Value
        b = Workbooks(Dateiname).Worksheets(3).Cells(8, 6).Value
    'End If
    
    fixeduser = "??"
    usercount = 1  'wenigstens einen user holen, auch wenn eintrag leer
    For i = 2 To 50
        If Workbooks(Dateiname).Worksheets(3).Cells(i, 1).Value = "" Then
            usercount = i - 2
            i = 51
        End If
    Next i
    usercount = Workbooks(Dateiname).Worksheets(3).UsedRange.Rows.Count - 1
    'nutzer im array anlegen
    For i = 0 To usercount - 1
        usernames(i) = Workbooks(Dateiname).Worksheets(3).Cells(i + 2, 1).Value
        
        If Workbooks(Dateiname).Worksheets(3).Cells(i + 2, 2).Value <> "" And fixeduser = "??" Then ' user setzen wenn kreuz
             fixeduser = usernames(i)
        End If
    Next i

    
End Sub

Public Sub holeDatenbank()
    On Error GoTo FehlerOpen
  '  If Workbooks(Dateiname).Worksheets(3).Cells(11, 4).Value = "x" Then
   '     GoTo FehlerOpen
   ' End If
    Workbooks.Open Filename:=a & b, ReadOnly:=True, Password:=pwlager
    On Error GoTo 0
    
    Workbooks(Dateiname).Worksheets(3).Cells(1, 6).Value = "Terminal Version " & Str(VERSION) 'versionsnummer aus konstante in terminal schreiben
    If IsNumeric(Workbooks(b).Worksheets(1).Cells(2, 22).Value) Then    ' auf versionsnummer der datenbank prüfen
        If VERSION < Workbooks(b).Worksheets(1).Cells(2, 22).Value Then
            MsgBox ("Für das Terminal gibt es ein Update. Bitte neue Version vom Server laden")
            Workbooks(b).Close SaveChanges:=False
            Workbooks(Dateiname).Close SaveChanges:=False
            Exit Sub
        End If
    Else
        MsgBox ("Für das Terminal gibt es ein Update. Bitte neue Version vom Server laden")
        Workbooks(b).Close SaveChanges:=False
        Workbooks(Dateiname).Close SaveChanges:=False
        
        Exit Sub
    End If
'     kopieren der arbeitsblätter
'    If Not init Then
'        'löschen der alten blätter
'        If Workbooks(Dateiname).Sheets.Count > 3 Then
'            Application.DisplayAlerts = False
'            For k = Workbooks(Dateiname).Sheets.Count To 4 Step -1
'                Workbooks(Dateiname).Worksheets(k).Delete
'            Next k
'            Application.DisplayAlerts = True
'        End If
'        'kopieren
'        If Workbooks(b).Sheets.Count > 1 Then
'            Dim Projektblatt As Worksheet
'            For k = 1 To Workbooks(b).Sheets.Count
'                Set Projektblatt = Workbooks(b).Worksheets(k)
'                Projektblatt.Copy After:=Workbooks(Dateiname).Worksheets(k + 2)
'            Next k
'        Else
'            MsgBox "keine Projektblätter gefunden"
'        End If
'        init = True
'        Workbooks(Dateiname).Worksheets(1).Activate
'        Workbooks(Dateiname).Save
'    End If
    
    ll = Workbooks(b).Worksheets(1).UsedRange                           ' inhalt Lagerliste als variant in RAM
    llrows = Workbooks(b).Worksheets(1).UsedRange.Rows.Count            'anzahl der zeilen ausgeben
    llcols = Workbooks(b).Worksheets(1).UsedRange.Columns.Count         'anzahl der spalten ausgeben
    
    plcount = Workbooks(b).Sheets.Count - 1 'beim plcount wird lager nicht mitgezähl
    If plcount > 50 Then
        MsgBox "maximal 50 Projektblätter zulässig"
    ElseIf plcount > 0 Then
       For i = 2 To (plcount + 1)
            pl(i - 2) = Workbooks(b).Worksheets(i).UsedRange                    ' inhalt pl als variant in RAM
            plrows(i - 2) = Workbooks(b).Worksheets(i).UsedRange.Rows.Count          'anzahl der zeilen ausgeben
            plnames(i - 2) = Workbooks(b).Worksheets(i).Name
        Next i
    Else
        MsgBox "keine Projektblätter gefunden"
    End If
    
    Workbooks(b).Close SaveChanges:=False
    
    Exit Sub
FehlerOpen:

    MsgBox ("Fehler beim Auslesen der Stammdaten vom Server. Bitte Dateipfad und Namen im tabellenblatt 'Einstellungen' prüfen!")
    'Workbooks(Dateiname).Worksheets(3).Cells(11, 4).Value = "x"
    
    
   ' ll = Workbooks(Dateiname).Worksheets(4).UsedRange                           ' inhalt Lagerliste als variant in RAM
  '  llrows = Workbooks(Dateiname).Worksheets(4).UsedRange.Rows.Count            'anzahl der zeilen ausgeben
   ' llcols = Workbooks(Dateiname).Worksheets(4).UsedRange.Columns.Count         'anzahl der spalten ausgeben
    
  '  plcount = Workbooks(Dateiname).Sheets.Count - 4
  '  If plcount > 50 Then
   '     MsgBox "maximal 50 Projektblätter zulässig"
   ' ElseIf plcount > 1 Then
   '    For i = 5 To plcount
   '         pl(i - 5) = Workbooks(Dateiname).Worksheets(i).UsedRange                    ' inhalt pl als variant in RAM
   '         plrows(i - 5) = Workbooks(Dateiname).Worksheets(i).UsedRange.Rows.Count          'anzahl der zeilen ausgeben
   '         plnames(i - 5) = Workbooks(Dateiname).Worksheets(i).Name
    '    Next i
   ' Else
    '    MsgBox "keine Projektblätter gefunden"
   ' End If
End Sub

Public Sub BatchBuchenÖffnen()
   ' Call Inits
    Call holeDatenbank
    'nutzer anlegen

    For k = 0 To usercount - 1
        BatchBuchungen.Nutzer.AddItem usernames(k)
        
        If usernames(k) = fixeduser Then
            BatchBuchungen.Nutzer.ListIndex = k
            Exit For
        End If
    Next k

    For k = 0 To plcount
        BatchBuchungen.Projektauswahl.AddItem plnames(k)
    Next k
    BatchBuchungen.Projektauswahl.ListIndex = -1 'erzwingen der auswahl
    
    BatchBuchungen.Show
    
End Sub

Sub checkInventurmodus()
    Dim i As Integer
    Inventurmodus = False
    Buchungsterminal.TextInventurmodus.Visible = False
    
    If Workbooks(Dateiname).Worksheets(3).Cells(12, 4).Value <> "" Then               'Inventurmodus angehakt
        Dim gefunden As Boolean
        gefunden = False
        Dim hlpstr1 As String
        Dim hlpstr2 As String
        hlpstr2 = "Inventur"
        If plcount > 0 Then 'automatisch auf projektblatt setzen welches stichwort "inventur" enthält
            For i = 0 To plcount - 1
                If InStr(1, plnames(i), hlpstr2, 1) > 0 Then
                    gefunden = True
                    Buchungsterminal.Projektauswahl.ListIndex = i + 1
                    Exit For
                End If
            Next i
        End If
        If Not gefunden Then
            MsgBox ("Inventurmodus nicht möglich. Bitte Projektblatt mit Stichwort -Inventur- im Namen anlegen")
            
        Else
            Buchungsterminal.ToggleBedarf.Visible = False
            Buchungsterminal.Einkaufsbuchung.Visible = False
            Buchungsterminal.ToggleBedarf.Value = False
            Buchungsterminal.ToggleBestand.Value = True
            Inventurmodus = True
            Buchungsterminal.TextInventurmodus.Visible = True
        End If
    Else
        Buchungsterminal.ToggleBedarf.Visible = True
    End If
    
    
End Sub

Sub ÖffneTerminal()

   ' Call Inits
    Call holeDatenbank
    
    If userlevel = 0 Then      'nutzergruppenverwaltung
        Buchungsterminal.EditScannen.Enabled = False
        Buchungsterminal.Einkaufsbuchung.Visible = False
        Buchungsterminal.Einkaufsbuchung.Value = False
    Else
    
    End If
    
    
    For k = 0 To usercount - 1
        Buchungsterminal.Nutzer.AddItem usernames(k)
        
        If usernames(k) = fixeduser Then
            Buchungsterminal.Nutzer.ListIndex = k
            Exit For
        End If
    Next k
    
    
    Buchungsterminal.Projektauswahl.Clear
    Buchungsterminal.Projektauswahl.AddItem "Lager"
    

    If plcount > 0 Then
        Buchungsterminal.Projektauswahl.ListIndex = 0
        For i = 0 To plcount - 1
            Buchungsterminal.Projektauswahl.AddItem plnames(i) 'projektauswahlliste füllen
        Next i
    Else
        MsgBox "keine Projektblätter gefunden"
    End If
    Buchungsterminal.Projektauswahl.ListIndex = -1 'erzwingen der auswahl

    Call EinträgeLöschen
    
    Buchungsterminal.Scancode.SetFocus
    Buchungsterminal.nBuchungen = 1
    letzteGültigeZahl = 1
    Buchungsterminal.Nachbestellen.Visible = False
    Buchungsterminal.ToggleBedarf.Value = False
    Buchungsterminal.ToggleBestand.Value = True
    Buchungsterminal.Wann.Visible = False
    Buchungsterminal.HLWann.Visible = False
    Buchungsterminal.HintWann.Visible = False
    scan = True
    
    Call checkInventurmodus
       
    Buchungsterminal.Show
  
End Sub

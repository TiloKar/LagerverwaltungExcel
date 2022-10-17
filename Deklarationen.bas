Attribute VB_Name = "Deklarationen"


Public Const VERSION As Integer = 11
Public b As String '= "Lagerliste 2017.xlsx"                   'dateiname lagerliste
Public Const c As String = "Journal.xlsx"                   'dateiname lagerliste
Public Const d As String = "Problemkinder.xlsx"                   'dateiname lagerliste
Public Const pwjournal As String = "journal"
Public Const pwlager As String = "lager"
                           'Eintrag hällt user auf diesem wert
Public Const nachkomma As Integer = 1

Public Const maxRowsSLP = 200

'Public a As String
Public ll As Variant    ' Ram image lagerliste
Public llrows As Integer
Public llcols As Integer

Public pl(50) As Variant    ' Ram image projektlisten
Public plrows(50) As Integer
Public plcount As Integer
Public plnames(50) As String

Public usernames(50) As String
Public fixeduser As String
Public usercount As Integer

Public userlevel As Integer
Public BedarfVorBuchung As Integer
Public Inventurmodus As Boolean
'Public init As Boolean
Public a As String
Public rpZeile As Integer
Public Z As Integer
Public S As Integer
Public keinPiep As Boolean
Public scan As Boolean
Public letzteGültigeZahl As Single
Public EsWurdeAufButtonGescannt As Boolean
Public Dateiname As String
Public strListe As Variant
Public strPS As Variant
'Declare Function Beep Lib "kernel32.dll" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Sucheigenschaft As Integer
Public EigenschaftGeändert As Boolean

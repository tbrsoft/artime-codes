Attribute VB_Name = "Module1"
Public USR As String
Public PSW As String
Public PSWm1 As String
Public PSWm2 As String
Public CN As New ADODB.Connection

'------------------------- PAVADAS --------------------------------------
Public Sub PintarTxt(TXT As Control)
    TXT.SelStart = 0
    TXT.SelLength = Len(TXT)
End Sub


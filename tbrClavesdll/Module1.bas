Attribute VB_Name = "Module1"
Public AP As String
Public KriP As New tbrCrypto.Crypt

Public Function stFechaSQL(FECHA As Date) As String
    Dim FechaChota As String    'sql tiene la fecha al reves por eso
    Dim h() As String
    
    h = Split(CStr(FECHA), "/")
    FechaChota = h(1) + "/" + h(0) + "/" + h(2)
    
    stFechaSQL = FechaChota
End Function

Public Function NoNuloN(J) As Single
    If IsNumeric(J) Then
        NoNuloN = J
    Else
        NoNuloN = 0
    End If
End Function

Public Function NoNuloS(S) As String
    If IsNull(S) Then
        NoNuloS = ""
    Else
        NoNuloS = S
    End If
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTbrNumeroLetra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function NumeroenLetras(Numero As String) As String
    Dim sNumero() As String, tmpEntero As String, tmpDecimal As String
    'supongo que siempre es numerico el nro
    'Numero = CStr(CSng(Numero)) 'por si tiene separador de miles
    'no anda por el exponencial
    
    '(1) lo paso a string y separo lo decimal de lo entero
    If InStrRev(Numero, ",") = 0 Then
        sNumero = Split(Numero, ".")
    Else
        sNumero = Split(Numero, ",")
    End If
    
    '¿si no tiene decimales?, le agrego uno con 0
    If UBound(sNumero) = 0 Then
        ReDim Preserve sNumero(1)
        sNumero(1) = "0"
    End If
    
    'ya lo tenemos string
    '(2)ENTERO!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Dim xX As Long
    xX = Len(sNumero(0))
    tmpEntero = ""
    
    Select Case xX
        Case 1 '0-9
            tmpEntero = NumeritoLetra(sNumero(0))
        Case 2 '0-99
            tmpEntero = DosCifras(sNumero(0))
        Case 3 '0-999
            tmpEntero = TresCifras(sNumero(0))
        Case 4 '0-9999
            tmpEntero = CuatroCifras(sNumero(0))
        Case 5 '0-99.999
            tmpEntero = CincoCifras(sNumero(0))
        Case 6 '0-999.999
            tmpEntero = SeisCifras(sNumero(0))
        Case Is > 6 'millones
            tmpEntero = Millones(sNumero(0))
        Case 0
            tmpEntero = "cero"
    End Select
    
    '(3)DECIMALES!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Dim Lx As Long, nR As String
    Lx = Len(sNumero(1))
    
    If Lx <= 2 Then
        nR = Mid(sNumero(1), 1, Lx) + String(2 - Lx, "0")
    Else
        If CLng(Mid(sNumero(1), 3, 1)) < 5 Then 'redondeo
            nR = Mid(sNumero(1), 1, 2)
        Else
            nR = CStr(CLng(Mid(sNumero(1), 1, 2)) + 1)
        End If
    End If
    
    If nR = "" Or nR = "00" Or nR = "0" Then
        tmpDecimal = ""
    Else
        tmpDecimal = nR + "/100"
    End If
    
    'ultimo retoque
    tmpEntero = Replace(tmpEntero, "uno mil", "un mil")
    
    NumeroenLetras = tmpEntero + " " + tmpDecimal
End Function

Private Function NumeritoLetra(Numero As String) As String
    'recibe el numero 0-9 y lo escribe en letra
    Select Case Numero
        Case "0"
            NumeritoLetra = "cero"
        Case "1"
            NumeritoLetra = "uno"
        Case "2"
            NumeritoLetra = "dos"
        Case "3"
            NumeritoLetra = "tres"
        Case "4"
            NumeritoLetra = "cuatro"
        Case "5"
            NumeritoLetra = "cinco"
        Case "6"
            NumeritoLetra = "seis"
        Case "7"
            NumeritoLetra = "siete"
        Case "8"
            NumeritoLetra = "ocho"
        Case "9"
            NumeritoLetra = "nueve"
        Case Else
            NumeritoLetra = "mal"
    End Select
End Function

Private Function DosCifras(Numero As String) As String
    Dim Uno As Long, Dos As Long
    
    Uno = CLng(Mid(Numero, 1, 1))
    Dos = CLng(Mid(Numero, 2, 1))
    
    Select Case Uno
        Case 1
            Select Case Dos
                Case 0
                    DosCifras = "diez"
                Case 1
                    DosCifras = "once"
                Case 2
                    DosCifras = "doce"
                Case 3
                    DosCifras = "trece"
                Case 4
                    DosCifras = "catorce"
                Case 5
                    DosCifras = "quince"
                Case 6 To 9
                    DosCifras = "dieci" + NumeritoLetra(CStr(Dos))
                Case Else
                    DosCifras = "mal"
            End Select
        
        Case 2 'veinti...
            If Dos = 0 Then
                DosCifras = "veinte"
            Else
                DosCifras = "veinti" + NumeritoLetra(CStr(Dos))
            End If
            
        Case 3 'treinta y
            DosCifras = "treinta"
            If Dos <> 0 Then DosCifras = DosCifras + " y " + NumeritoLetra(CStr(Dos))
            
        Case 4
            DosCifras = "cuarenta"
            If Dos <> 0 Then DosCifras = DosCifras + " y " + NumeritoLetra(CStr(Dos))
            
        Case 5
            DosCifras = "cincuenta"
            If Dos <> 0 Then DosCifras = DosCifras + " y " + NumeritoLetra(CStr(Dos))
        Case 6
            DosCifras = "sesenta"
            If Dos <> 0 Then DosCifras = DosCifras + " y " + NumeritoLetra(CStr(Dos))
        Case 7
            DosCifras = "setenta"
            If Dos <> 0 Then DosCifras = DosCifras + " y " + NumeritoLetra(CStr(Dos))
        Case 8
            DosCifras = "ochenta"
            If Dos <> 0 Then DosCifras = DosCifras + " y " + NumeritoLetra(CStr(Dos))
        Case 9
            DosCifras = "noventa"
            If Dos <> 0 Then DosCifras = DosCifras + " y " + NumeritoLetra(CStr(Dos))
        Case 0
            DosCifras = NumeritoLetra(CStr(Dos))
            
    End Select
End Function

Private Function TresCifras(Numero As String) As String
    Dim Uno As Long, Dos As String
    
    Uno = CLng(Mid(Numero, 1, 1))
    Dos = Mid(Numero, 2, 2) 'tiene las 2 ultimas cifras
    
    Select Case Uno
        Case 1
            TresCifras = "cien"
            If Dos <> "00" Then TresCifras = TresCifras + "to " + DosCifras(Dos)
        Case 5
            TresCifras = "quinientos"
            If Dos <> "00" Then TresCifras = TresCifras + " " + DosCifras(Dos)
        Case 7
            TresCifras = "setecientos"
            If Dos <> "00" Then TresCifras = TresCifras + " " + DosCifras(Dos)
        Case 9
            TresCifras = "novecientos"
            If Dos <> "00" Then TresCifras = TresCifras + " " + DosCifras(Dos)
        Case 0
            TresCifras = DosCifras(Dos)
        Case 2, 3, 4, 6, 8
            TresCifras = NumeritoLetra(CStr(Uno)) + "cientos"
            If Dos <> "00" Then TresCifras = TresCifras + " " + DosCifras(Dos)
    End Select
End Function

Private Function CuatroCifras(Numero As String) As String
    Dim Uno As String, Dos As String
    
    Uno = Mid(Numero, 1, 1)
    Dos = Mid(Numero, 2, 3) 'tiene las 3 ultimas cifras
    
    If Dos = "000" Then
        If Uno = "1" Then
            CuatroCifras = "mil"
        Else
            CuatroCifras = NumeritoLetra(Uno) + " mil"
        End If
    Else
        Select Case Uno
            Case 0
                CuatroCifras = TresCifras(Dos)
            Case 1
                CuatroCifras = "mil " + TresCifras(Dos)
            Case Else
                CuatroCifras = NumeritoLetra(Uno) + " mil " + TresCifras(Dos)
        End Select
    End If
End Function

Private Function CincoCifras(Numero As String) As String
    Dim Uno As String, Dos As String
    
    Uno = Mid(Numero, 1, 2) 'tiene las 2 primeras cifras
    Dos = Mid(Numero, 3, 3) 'tiene las 3 ultimas cifras
    
    If Left(Numero, 1) = "0" Then
        CincoCifras = CuatroCifras(Right(Numero, 4))
        Exit Function
    End If
    
    CincoCifras = DosCifras(Uno) + " mil "
    
    If Dos <> "000" Then CincoCifras = CincoCifras + TresCifras(Dos)
    
End Function

Private Function SeisCifras(Numero As String) As String
    Dim Uno As String, Dos As String
    
    Uno = Mid(Numero, 1, 3) 'tiene las 3 primeras cifras
    Dos = Mid(Numero, 4, 3) 'tiene las 3 ultimas cifras
    
    If Left(Numero, 1) = "0" Then
        SeisCifras = CincoCifras(Right(Numero, 5))
        Exit Function
    End If
    
    SeisCifras = TresCifras(Uno) + " mil "
    
    If Dos <> "000" Then SeisCifras = SeisCifras + TresCifras(Dos)
    
End Function

Private Function Millones(Numero As String) As String
    Dim lenM As Long, Uno As String, Dos As String
    
    lenM = Len(Numero) - 6
    
    If lenM <= 6 Then 'hasta 999 mil millones
        Uno = String(6 - lenM, "0") + Mid(Numero, 1, lenM)
        Dos = Mid(Numero, lenM + 1, 6)
       
        Millones = SeisCifras(Uno) + " millones "
        If Dos <> "000000" Then Millones = Millones + SeisCifras(Dos)
    Else 'mas no hace
        Millones = "Numero muy grande"
    End If
End Function

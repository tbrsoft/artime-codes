VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mArchMDB As String
Private mArchMDB2 As String
Private mCN As New ADODB.Connection
Private CN2 As New ADODB.Connection
Private mPsW As String
Private mPsW2 As String
Dim Datos() As String
Dim FSO As New FileSystemObject
Dim WF As String

Public Property Let PSW2(newPSW2 As String)
    mPsW2 = newPSW2
End Property

Public Property Get ArchMDB2() As String
    ArchMDB2 = mArchMDB2
End Property

Public Property Let ArchMDB2(NewArch2 As String)
    mArchMDB2 = NewArch2
End Property

Public Property Let PSW(newPSW As String)
    mPsW = newPSW
End Property

Public Property Get ArchMDB() As String
    ArchMDB = mArchMDB
End Property

Public Property Let ArchMDB(NewArch As String)
    mArchMDB = NewArch
End Property

Public Sub Conectar()
    mCN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _
        mArchMDB + ";Jet OLEDB:Database Password=" + mPsW
    mCN.Open
End Sub

Public Sub ConectarCliente()
    CN2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _
        mArchMDB2 + ";Jet OLEDB:Database Password=" + mPsW2
    CN2.Open
End Sub

Public Sub CN_CLOSE()
    If mCN.State = adStateOpen Then mCN.Close
    If CN2.State = adStateOpen Then CN2.Close
End Sub

Private Sub Class_Initialize()
    WF = FSO.GetSpecialFolder(WindowsFolder)
    If Right(WF, 1) <> "\" Then WF = WF + "\"
End Sub

Private Sub Class_Terminate()
    If mCN.State = adStateOpen Then mCN.Close
    If CN2.State = adStateOpen Then CN2.Close
    Set FSO = Nothing
End Sub

Public Function GetIDSiesNuevo(CUUsuario As String, _
    Optional Nombre As String = "SIN NOMBRE TODAVIA") As String
    Dim CUUc As String
    
    If mCN.State = adStateClosed Then Conectar
    
    If CUUsuario = "111111" Then
        CUUc = CStr(AgregarCliente(Nombre))
        mCN.Execute "UPDATE Claves SET CUUsuario = " + CUUc + _
            " WHERE CUUsuario = 111111"
    Else
        CUUc = CUUsuario
    End If
    
    GetIDSiesNuevo = CUUc
End Function

Public Function GenerarClave(CUUsuario As String, Creditos As Long, ClaveDK As String) As String
    'el archivo esta en wf y es "bio2.enb"
    Dim CUUc As String, Clave As String, ClaveDK2 As String
    
    CUUc = CUUsuario
    Datos = Split(ClaveDK, "|")
    Datos(1) = CUUc
    ClaveDK2 = Join(Datos, "|")
    '---------------------------------------------------------------------------
    
    'genero clave de esta manera se tiene que corresponder con el pedido---------
    'la clave es toda numerica---------------------------------------------------
    
    'todo va a ir sumado por la primera cifra (si es 3 al 8 lo transforma en 1)
    'son en total 24 cifras
    '(FechaUsuarioAlqTAlqP (23 cifras)
    'como son: (0=7,1=3,2=9,3=1,4=8,5=0,6=6,7=2,8=5,9=4)
    'fecha: MesAno -> mes: 2 cifras
'    (3 cifras)    -> año: ultima cifra
    'usuario:7 cifras
    'alqt:ult 3 cifras
    'alqp:ult 3 cifrs
    'creditos: 7 cifras completadas)
    'creditos segun configuracion de precios
        
    'empiezo poniendo los nros originales(sin el primero)
    Clave = OrigClaveOriginal(Creditos)
    
    'cambio las cifras segun el parentesis
    Clave = ClaveParentesis(Clave)
    
    'agrego 1 cifra(aleatorio)-sumo a c/cifra ese valor (quedan 20 cifras)
    Clave = ClaveDefinitivA(Clave)
    '----------------------------------------------------------------------------
   
    GenerarClave = Clave
End Function

Private Function ClaveDefinitivA(Clave As String) As String
    Dim ALeat As Long, I As Long, Tmp As Long, tmP2 As String, eXN As Long
    Randomize   ' Inicializa el generador de números aleatorios.
    
    ALeat = Int((9 * Rnd))    ' Genera valores aleatorios entre 0 y 9.
    tmP2 = CStr(ALeat)
    
    For I = 1 To Len(Clave)
        eXN = CLng(Mid(Clave, I, 1))
        Tmp = eXN + ALeat
        
        If Tmp > 9 Then 'es de 2 cifra
            tmP2 = tmP2 + Right(CStr(Tmp), 1)
        Else
            tmP2 = tmP2 + CStr(Tmp)
        End If
        
    Next I
    
    ClaveDefinitivA = tmP2
End Function

Private Function ClaveParentesis(ExClave As String) As String
    'como son: (0=7,1=3,2=9,3=1,4=8,5=0,6=6,7=2,8=5,9=4)
    Dim tmpClave As String, I As Long, eXN As String
    
    tmpClave = ""
    For I = 1 To Len(ExClave)
        eXN = Mid(ExClave, I, 1)
        
        Select Case eXN
            Case "0"
                tmpClave = tmpClave + "7"
            Case "1"
                tmpClave = tmpClave + "3"
            Case "2"
                tmpClave = tmpClave + "9"
            Case "3"
                tmpClave = tmpClave + "1"
            Case "4"
                tmpClave = tmpClave + "8"
            Case "5"
                tmpClave = tmpClave + "0"
            Case "6"
                tmpClave = tmpClave + "6"
            Case "7"
                tmpClave = tmpClave + "2"
            Case "8"
                tmpClave = tmpClave + "5"
            Case "9"
                tmpClave = tmpClave + "4"
        End Select
            
    Next I
    
    ClaveParentesis = tmpClave
    
End Function

Private Function OrigClaveOriginal(Creditos As Long) As String
    'empiezo poniendo los nros originales(sin el primero)
    Dim FH As String, US As String, AT As String, AD As String, CR As String
    Dim FECHA() As String
    
    FECHA = Split(Datos(0), "/") 'mes - dia - año
                'mes (2 cifras)
    FH = String(2 - Len(FECHA(0)), "0") + FECHA(0) + _
        Right(FECHA(2), 1) 'año (ultima cifra)
    US = String(7 - Len(Datos(1)), "0") + Datos(1) 'son 7 cifras
    If Len(Datos(3)) <= 3 Then
        AT = String(3 - Len(Datos(3)), "0") + Datos(3)
    Else
        AT = Right(Datos(3), 3)
    End If
        
    If Len(Datos(4)) <= 3 Then
        AD = String(3 - Len(Datos(4)), "0") + Datos(4)
    Else
        AD = Right(Datos(4), 3)
    End If
    
    CR = String(7 - Len(CStr(Creditos)), "0") + CStr(Creditos)
    
    OrigClaveOriginal = FH + US + AT + AD + CR
    
End Function

Public Function AgregarCliente(Optional Usuario As String = "SIN NOMBRE TODAVIA") As Long
    'devuelve el CUU del cliente
    Dim Tmp As Long
    
    Tmp = NuevoCUU + 1
    mCN.Execute "INSERT INTO Usuarios (CUUsuario,Usuario) VALUES (" + _
        CStr(Tmp) + ",'" + Usuario + "')"
    
    AgregarCliente = Tmp
End Function

Private Function NuevoCUU() As Long
    Dim Tmp As Long
    Dim RsT As New ADODB.Recordset
    
    RsT.Open "SELECT TOP 1 * FROM Usuarios ORDER BY CUUsuario DESC", mCN, adOpenStatic, adLockReadOnly
    
    Tmp = CLng(RsT("CUUsuario"))
    
    RsT.Close
    Set RsT = Nothing
    
    NuevoCUU = Tmp
End Function

Private Function FileContar(FolderSpec As String) As Long
    Dim F, F1, Fc, S As Long
    
    Set F = FSO.GetFolder(FolderSpec)
    Set Fc = F.Files
    For Each F1 In Fc
        S = S + 1
    Next
    
    FileContar = S
End Function
    
Private Function ListarFiles(EspecificacionDeCarpeta As String) As String()
    Dim FSO As New FileSystemObject, F, F1, Fc, S As Long
    Dim Archs() As String
    
    Set F = FSO.GetFolder(EspecificacionDeCarpeta)
    Set Fc = F.Files
    
    ReDim Preserve Archs(0)
    Archs(0) = ""
    S = 1
    
    For Each F1 In Fc
      ReDim Preserve Archs(S)
      Archs(S) = F1.Name
      S = S + 1
    Next
    
    ListarFiles = Archs
    
    Set FSO = Nothing
End Function

Public Sub DeKripArch(Archivo As String, SuAP As String)
    'los .enb estan encriptados, los .rec decriptados
    'solo los .rec ya se cargaron a la base de datos
    Dim Arch2 As String, TE As TextStream, Tmp As String, SP() As String
    Dim NArchivo As String, I As Long
    
    NArchivo = Right(Archivo, Len(Archivo) - InStrRev(Archivo, "\"))
    Arch2 = SuAP + "Pedidos\sinconexion.rec"
    
    KriP.DecryptFile eMC_Skipjack, Archivo, Arch2, "siboldi"
    
    FSO.DeleteFile Archivo
    'lo grabo en la BD
    Set TE = FSO.OpenTextFile(Arch2, ForReading, True)
    Tmp = TE.ReadLine
    TE.Close
    
    SP = Split(Tmp, "|") '(fecha ya esta SQL)
    
    For I = 0 To UBound(SP)
        SP(I) = NoNuloS(SP(I))
    Next I
    
    mCN.Execute "INSERT INTO Claves (Fecha, ClaveDK, Dolares, " + _
        "CUUsuario, Archivo, Enviado) VALUES (#" + SP(0) + "#,'" + _
        Tmp + "'," + Replace(SP(2), ",", ".") + "," + SP(1) + ",'" + _
        NArchivo + "','NO')"
    FSO.DeleteFile Arch2
End Sub

Public Sub RecibirPedidos()
    'pasa los .enb a .rec y los decripta
    Dim TE As TextStream, Tmp As String, tmP2 As String
    Dim ArchivoMDB As String, Arch As String, Arch2 As String
    Dim Recib() As String, J As Long
    
    AP = Left(mArchMDB, InStrRev(mArchMDB, "\"))
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    Recib = ListarFiles(AP + "Pedidos\Archivos Bajados\")
    If UBound(Recib) = 0 Then Exit Sub
        
    For I = 1 To UBound(Recib)
        'los .enb estan encriptados, los .rec decriptados
        'solo los .rec ya se cargaron a la base de datos
        Arch = AP + "Pedidos\Archivos Bajados\" + Recib(I)
        tmP2 = Replace(Recib(I), "enb", "rec")
        Arch2 = AP + "Pedidos\" + tmP2
        
        KriP.DecryptFile eMC_Skipjack, Arch, Arch2, "siboldi"
        
        FSO.DeleteFile Arch
        'lo grabo en la BD
        Set TE = FSO.OpenTextFile(Arch2, ForReading, True)
        Tmp = TE.ReadLine
        TE.Close
        
        SP = Split(Tmp, "|") '(fecha ya esta SQL)
        For J = 0 To UBound(SP)
            SP(J) = NoNuloS(SP(J))
        Next J
        
        mCN.Execute "INSERT INTO Claves (Fecha, ClaveDK, Dolares, " + _
            "CUUsuario, Archivo, Enviado) VALUES (#" + SP(0) + "#,'" + _
            Tmp + "'," + Replace(SP(2), ",", ".") + "," + SP(1) + ",'" + _
            tmP2 + "','NO')"
        FSO.DeleteFile Arch2
        
    Next I
    
    If mCN.State = adStateOpen Then mCN.Close
    
    Conectar
End Sub

Public Sub EnviarPedido(IdUser As Long, Dol As Single, AlqT As Long, AlqP As Long, _
    NumArch As Long)
    
    Dim TE As TextStream
    
    ' pongo en la base de datos al pedido y al archivo a ser enviado
    'base de datos---------------------------------------------------------------
    CN2.Execute "INSERT INTO Estadisticas (Fecha,Compras,Pedidos,Deudas," + _
        "Comentarios,ID,IDGral) VALUES (#" + stFechaSQL(Date) + "#," + _
        Replace(CStr(Dol), ",", ".") + "," + CStr(AlqT) + _
        "," + CStr(AlqP) + ",'NO'," + CStr(AlqT) + "," + CStr(NumArch) + ")"
    
    'en el archivo (separados por "|")-------------------------------------------
    Tmp = stFechaSQL(Date) + "|" + CStr(IdUser) + "|" + _
        CStr(Dol) + "|" + CStr(AlqT) + "|" + CStr(AlqP)
    
    If FSO.FileExists(WF + "bio.enb") = False Then
        Set TE = FSO.CreateTextFile(WF + "bio.enb")
    Else
        Set TE = FSO.OpenTextFile(WF + "bio.enb", ForWriting)
    End If
    
    TE.WriteLine Tmp
    TE.Close
    
    'encripto archivo
    
    KriP.EncryptFile eMC_Skipjack, WF + "bio.enb", WF + "bio2.enb", "siboldi"
    'borro el original
    FSO.DeleteFile WF + "bio.enb"
       
End Sub

Public Function TienePedidos() As Boolean
    Dim RsT As New ADODB.Recordset
    
    'veo si tiene hecho algun pedido
    RsT.Open "SELECT * FROM Estadisticas WHERE Comentarios = 'NO'", CN2, adOpenStatic, adLockReadOnly
    'hay como maximo 1 pedido
    
    If RsT.RecordCount = 0 Then
        RsT.Close
        Set RsT = Nothing
        TienePedidos = False
        Exit Function
    End If
    
    TienePedidos = True
    
    RsT.Close
    Set RsT = Nothing
End Function

Public Function IngresoContrasena(Contrasena As String, AlqPendientes As Long) As Long
    '-1 mal
    '-2 los alq pendientes que tenia cuando compro son < a los que tiene ahora
        ' (algo raro hizo)
    '-3 es otro usuario
    '>=0 OK son las cant de creditos que se cargaron
    
    Dim Tmp As String, KlaweLeida As String, tmPN As Long, I As Long, eXN As Long
    Dim CUUs As Long, Kreditos As Long
    Dim RSQ As New ADODB.Recordset
    Dim RsT As New ADODB.Recordset
    
    RsT.Open "SELECT * FROM Estadisticas WHERE Comentarios = 'NO'", _
        CN2, adOpenStatic, adLockReadOnly
       
    Tmp = Contrasena
    
    'leo la clave---------------------------------------------------------------
    '1 saco el aleatorio y los resto
    tmPN = CLng(Left(Tmp, 1))
    Tmp = Right(Tmp, Len(Tmp) - 1)
    KlaweLeida = ""
    
    For I = 1 To Len(Tmp)
        eXN = CLng(Mid(Tmp, I, 1))
        If eXN < tmPN Then
            KlaweLeida = KlaweLeida + CStr(10 + eXN - tmPN)
        Else
            KlaweLeida = KlaweLeida + CStr(eXN - tmPN)
        End If
    Next I
    
    '2 el parentesis
    KlaweLeida = KlaweParentesis(KlaweLeida)
        
        'datos para controlar (las de 2 son las que estan como pedido)
    Dim FH As String, US As String, AT As String, AD As String, CR As String
    Dim FH2 As String, US2 As String, AT2 As String, AD2 As String
    
    FH = Mid(KlaweLeida, 1, 3)
    US = Mid(KlaweLeida, 4, 7)
    AT = Mid(KlaweLeida, 11, 3)
    AD = Mid(KlaweLeida, 14, 3)
    CR = Mid(KlaweLeida, 17, 7)
    '----------------------------------------------------------------------------
    
    'comparo la clave ----------------------------------------------------------
    '1 veo si era usuario nuevo le pongo el nuevo cuus
    RSQ.Open "SELECT TOP 1 * FROM CodigoUsuario WHERE CUUsuario<>0", CN2, adOpenStatic, adLockReadOnly
    
    CUUs = CLng(RSQ("CUUsuario"))

    If CUUs = 111111 Then
        CN2.Execute "UPDATE CodigoUsuario SET CUUsuario = " + US + _
            " WHERE CUUsuario = 111111"
    Else
        '2 veo que el usuario sea para el que venga la clave
        Dim RSN As New ADODB.Recordset
        
        RSN.Open "SELECT TOP 1 CUUsuario FROM CodigoUsuario WHERE CUUsuario > 1", _
            CN2, adOpenStatic, adLockReadOnly
        US2 = CStr(RSN("CUUsuario"))
        
        RSN.Close
        Set RSN = Nothing
        
        If US <> US2 Then
            IngresoContrasena = -1
        End If
    End If
      
    '3 veo si estaba registrado el pedido
    FH2 = String(2 - Len(CStr(Month(CDate(RsT("Fecha"))))), "0") + _
        CStr(Month(CDate(RsT("Fecha")))) + _
        Right(CStr(Year(CDate(RsT("Fecha")))), 1)
    
    If Len(CStr(RsT("Pedidos"))) <= 3 Then
        AT2 = String(3 - Len(CStr(RsT("Pedidos"))), "0") + CStr(RsT("Pedidos"))
    Else
        AT2 = Right(CStr(RsT("Pedidos")), 3)
    End If
        
    If Len(CStr(RsT("Deudas"))) <= 3 Then
        AD2 = String(3 - Len(CStr(RsT("Deudas"))), "0") + CStr(RsT("Deudas"))
    Else
        AD2 = Right(CStr(RsT("Deudas")), 3)
    End If
    
    If FH = FH2 And AT = AT2 And AD = AD2 Then 'esta todo mas que joya
        'veo que entre el pedido y la compra no haya sumado creditos
            '(alqP0 < alqP1 algo hizo pero ilegalmente)
        If AlqPendientes > CLng(RsT("Deudas")) Then
            IngresoContrasena = -2
        Else
            'le acredito nomas
            IngresoContrasena = CLng(CR)
            
            'le marco que uso la clave
            CN2.Execute "UPDATE Estadisticas SET Comentarios = 'SI' " + _
                "WHERE Comentarios = 'NO'"
        End If
    Else        'algo esta mal
        IngresoContrasena = -1
    End If
    
    RsT.Close
    Set RsT = Nothing
    RSQ.Close
    Set RSQ = Nothing
    
End Function

Private Function KlaweParentesis(ExKlawe As String) As String
    'como son: (0=7,1=3,2=9,3=1,4=8,5=0,6=6,7=2,8=5,9=4) 'LOS REVIERTO
    Dim tmpClave As String, I As Long, eXN As String
    
    tmpClave = ""
    For I = 1 To Len(ExKlawe)
        eXN = Mid(ExKlawe, I, 1)
        
        Select Case eXN
            Case "7"
                tmpClave = tmpClave + "0"
            Case "3"
                tmpClave = tmpClave + "1"
            Case "9"
                tmpClave = tmpClave + "2"
            Case "1"
                tmpClave = tmpClave + "3"
            Case "8"
                tmpClave = tmpClave + "4"
            Case "0"
                tmpClave = tmpClave + "5"
            Case "6"
                tmpClave = tmpClave + "6"
            Case "2"
                tmpClave = tmpClave + "7"
            Case "5"
                tmpClave = tmpClave + "8"
            Case "4"
                tmpClave = tmpClave + "9"
        End Select
            
    Next I
    
    KlaweParentesis = tmpClave
    
End Function

Public Sub BorrarPedidos()
    'no los borrar los dejo como XX
    CN2.Execute "UPDATE Estadisticas SET Comentarios = 'XX' " + _
        "WHERE Comentarios = 'NO'"
    'ver que tbrSoft se entere de las cancelaciones para no responder
End Sub

Public Sub AcomodarInfoUsuario()
    Dim TE As TextStream, Recib() As String, I As Long, SP() As String
    Dim CUUsuario As String
    
    AP = Left(mArchMDB, InStrRev(mArchMDB, "\"))
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    Recib = ListarFiles(AP + "Infos\")
    If UBound(Recib) = 0 Then Exit Sub
        
    For I = 1 To UBound(Recib)
        'los .enb estan encriptados, los .rec decriptados
        'solo los .rec ya se cargaron a la base de datos
        Arch = AP + "Infos\" + Recib(I)
        
        'lo grabo en la BD
        Set TE = FSO.OpenTextFile(Arch, ForReading, True)
        Tmp = TE.ReadLine
        TE.Close
        
        SP = Split(Tmp, "'")
        If UBound(SP) < 5 Then
            Tmp = " ' ' ' ' "
            SP = Split(Tmp, "'")
        End If
        
        CUUsuario = SP(0)
        mCN.Execute "UPDATE Usuarios SET Nombre = '" + SP(1) + "'," + _
            "Direccion = '" + SP(2) + "'," + _
            "Telefono = '" + SP(3) + "'," + _
            "Pais = '" + SP(4) + "'," + _
            "Mail = '" + SP(5) + "' WHERE CUUsuario = " + CUUsuario
            
        FSO.DeleteFile Arch
      
    Next I
    
    If mCN.State = adStateOpen Then mCN.Close
    
    Conectar
    
End Sub




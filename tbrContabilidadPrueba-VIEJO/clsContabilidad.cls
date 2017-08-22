VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPruebaContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'este modulo permite crear y administrar un plan de cuenta
'incluye desde las losta de cuentas hasta los movimientos y balances
Private Enum TipoCuenta
    Activo = 1
    Pasivo = 2
    Perdida = 3
    Ganancia = 4
    PNeto = 5
End Enum

Private mArchMDB As String
Private mCN As New ADODB.Connection
Private mPsW As String

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
    
    ControlarSueldos 'para que no se vean en listaregresos
End Sub

Public Sub CN_CLOSE()
    If mCN.State = adStateOpen Then mCN.Close
End Sub

Public Function GetUltIDMasUno() As Long
    Dim RSU As New ADODB.Recordset
    
    RSU.Open "SELECT TOP 1 ID FROM tblCuentas ORDER BY ID DESC", mCN, adOpenStatic, adLockReadOnly
    
    GetUltIDMasUno = CLng(RSU("ID")) + 1
    
    RSU.Close
    Set RSU = Nothing
End Function

Public Function GetUltIDAsientoMasUno(Optional Libro As String = "LibroDiario") As Long
    Dim RSU As New ADODB.Recordset
    
    RSU.Open "SELECT TOP 1 IDAsiento FROM " + Libro + _
        " ORDER BY IDAsiento DESC", mCN, adOpenStatic, adLockReadOnly
    
    GetUltIDAsientoMasUno = CLng(RSU("IDAsiento")) + 1
    
    RSU.Close
    Set RSU = Nothing
End Function

Public Function AgregarCuenta(idCuenta As Long, _
    DependeDeCuenta As Long, Nombre As String, _
    Optional Descripcion As String = "") As Long
    '0 OK
    '1 datos vacios
    '2 ya existe cuenta
    '3 ya existe una cuenta con ese número
    
    If Nombre = "" Or DependeDeCuenta < 0 Then
        AgregarCuenta = 1
        Exit Function
    End If
    
    If ExisteNCuenta(Nombre) > 0 Then
        AgregarCuenta = 2
        Exit Function
    End If
    
    If GetNameCuenta(idCuenta) <> "NO EXISTE" Then
        AgregarCuenta = 3
        Exit Function
    End If
        
    Dim S As String
    S = "INSERT INTO tblCuentas (ID,IdAnt, Nombre, Descripcion, Modificable) " + _
        " VALUES (" + CStr(idCuenta) + "," + CStr(DependeDeCuenta) + _
        ",'" + Nombre + "','" + Descripcion + "',0)"

    
    mCN.Execute S
    AgregarCuenta = 0
End Function

Public Function ModificarCuenta(idCuenta As Long, Optional IDcuentaW As String = "Nada", _
    Optional Nombre As String = "Nada", Optional Descripcion As String = "Nada") As Long

    '0 OK
    '1 Existe cuenta con ese número
    '2 Existe cuenta con ese nombre
        
    Dim ID2 As Long, Nm2 As String, Desc2 As String
    
    '¿Existe otra cuenta con ese número?
    If IsNumeric(IDcuentaW) Then
        If idCuenta <> CLng(IDcuentaW) And GetNameCuenta(CLng(IDcuentaW)) <> "NO EXISTE" Then
            ModificarCuenta = 1
            Exit Function
        End If
    End If
    
    '¿Existe otra cuenta con ese nombre?
    If Nombre <> "Nada" And ExisteNCuenta(Nombre) > 0 Then
        'si no cambio el nombre no pasa nada
        If GetNameCuenta(idCuenta) <> Nombre Then
            ModificarCuenta = 2
            Exit Function
        End If
    End If
    
    If IDcuentaW = "Nada" Then
        ID2 = idCuenta
    Else
        ID2 = IDcuentaW
    End If
    
    If Nombre = "Nada" Then
        Nm2 = GetNameCuenta(idCuenta)
    Else
        Nm2 = Nombre
    End If
    
    If Descripcion = "Nada" Then
        Desc2 = ""
    Else
        Desc2 = ", Descripcion = '" + Descripcion + "'"
    End If
    
    S = "UPDATE tblCuentas SET Nombre = '" + _
        newNombre + "' WHERE Nombre = '" + tmpExNombre + "'"
    
    mCN.Execute "UPDATE tblCuentas SET ID = " + CStr(ID2) + ", " + _
        "Nombre = '" + Nm2 + "'" + Desc2 + " WHERE ID = " + CStr(idCuenta)
        
    'tengo que cambiar todos los IdAnterior
    If ID2 <> idCuenta Then
        mCN.Execute "UPDATE tblCuentas SET IDAnt = " + CStr(ID2) + " WHERE " + _
            "IDAnt = " + CStr(idCuenta)
    End If
    
    ModificarCuenta = 0
End Function

Public Sub EliminarCuenta(idCuenta As Long)
    Dim HIJ() As String, H As Long
    Dim SaldoConSub As Single, Nombre As String
    
    'primero cuanto tenia de saldo con subcuentas para mantener DEBE=HABER
    SaldoConSub = GetSaldo(idCuenta)
    Nombre = GetNameCuenta(idCuenta)
    
    '1ro que todo eliminar las hijas
    HIJ = GetCuentas(idCuenta)
    
    If UBound(HIJ) > 0 Then
        For H = 1 To UBound(HIJ)
            EliminarCuenta CLng(HIJ(H))
        Next H
    End If
    
    mCN.Execute "UPDATE LibroDiario SET IdCuenta = 57 WHERE IdCuenta = " + _
        CStr(idCuenta)
    mCN.Execute "UPDATE LibroSubDiario SET IdCuenta = 57 WHERE IdCuenta = " + _
        CStr(idCuenta)
    
    mCN.Execute "DELETE FROM tblCuentas WHERE ID = " + CStr(idCuenta)
End Sub

Public Sub GrabarCierre(IdCierre As Long, Efvo As Single, stResumen As String, _
    Dif As Single, Var As Single)
    'dejo grabado este cierre
    mCN.Execute "INSERT INTO CierresViejos (IDCierre,Fecha,Hora,Efvo,sCierre," + _
        "Diferencia,Variacion) " + _
        "VALUES (" + CStr(IdCierre) + ",#" + stFechaSQL(Date) + "#,'" + _
        CStr(Time) + "'," + _
        Replace(CStr(Efvo), ",", ".") + ",'" + stResumen + "'," + _
        Replace(CStr(Dif), ",", ".") + "," + _
        Replace(CStr(Var), ",", ".") + ")"
End Sub

Public Function GetCierre(IdCierre As Long) As String
    Dim rS As New ADODB.Recordset
    Dim stResumen As String
    
    rS.Open "SELECT sCierre FROM CierresViejos WHERE IdCierre = " + CStr(IdCierre), _
        mCN, adOpenStatic, adLockReadOnly
    
    If rS.RecordCount = 0 Then
        stResumen = "\\"
    Else
        stResumen = rS("sCierre")
    End If
    
    rS.Close
    Set rS = Nothing
    
    GetCierre = stResumen
    
End Function

Public Function GetRsCierres(Cantidad As Long) As ADODB.Recordset
    Dim rS3 As New ADODB.Recordset
    
    If rS3.State = adStateOpen Then rS3.Close
    
    rS3.CursorLocation = adUseClient
    rS3.Open "SELECT TOP " + CStr(Cantidad) + " * FROM CierresViejos ORDER BY ID DESC", _
        mCN, adOpenStatic, adLockReadOnly
    
    Set GetRsCierres = rS3
    'rS.Close
    'Set rS = Nothing
End Function

Public Function GetCuentas(IDBuscado As Long) As String()
    'devuelve todas las cuentas correspondientes a determinado Codigo de cuenta
    'devuelve los hijos de cada nivel. Con el parametro en cero devuelve los de la raiz
    
    'devuelve una matriz con los valores DESDE 1 en adelante si tiene solo cero
    'es por que no haya nada
    
    Dim Ret() As String
    ReDim Ret(0) 'para que devuelva algo por mas que no haya registros
    
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "Select * from tblCuentas Where IdAnt = " + CStr(IDBuscado), mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        GetCuentas = Ret
        Exit Function
    End If
    
    rS.MoveFirst
    Dim C As Long
    C = 1
    Do While Not rS.EOF
        ReDim Preserve Ret(C)
        Ret(C) = rS.Fields("id")
        rS.MoveNext
        C = C + 1
    Loop
    rS.Close
    Set rS = Nothing
    GetCuentas = Ret
    
End Function

Public Function GetNameCuenta(ID As Long) As String
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "SELECT * FROM tblCuentas WHERE Id = " + CStr(ID), mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        GetNameCuenta = "NO EXISTE"
        Exit Function
    End If
    
    GetNameCuenta = rS.Fields("nombre")
    rS.Close
    Set rS = Nothing
    
End Function

Public Function GetIDCuenta(Nombre As String) As Long
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "Select * from tblCuentas Where Nombre = '" + Nombre + "'", _
        mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        GetIDCuenta = -1
        Exit Function
    End If
    
    GetIDCuenta = rS.Fields("id")
    rS.Close
    Set rS = Nothing
    
End Function

Public Function GetDetalle(idCuenta As Long) As String
    Dim rS As New ADODB.Recordset
    
    rS.CursorLocation = adUseClient
    rS.Open "Select * from tblCuentas Where Id = " + CStr(idCuenta), _
        mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        GetDetalle = "NO EXISTE"
        Exit Function
    End If
    
    GetDetalle = NoNuloS(rS.Fields("Descripcion"))
    rS.Close
    Set rS = Nothing

End Function

Public Function EDITAR()
    frmEditarCuentas.Show 1
End Function

Private Sub Class_Terminate()
    If mCN.State = adStateOpen Then mCN.Close
End Sub

Private Sub AgregarMov(IdAsiento As Long, IdCta As Long, Var As Single, Fecha As Date, _
    Detalle As String, Optional Libro As String = "LibroDiario")
    
    mCN.Execute "INSERT INTO " + Libro + " (IdAsiento,IdCuenta,Variacion,Fecha,Detalle) " + _
        "VALUES (" + CStr(IdAsiento) + ", " + _
        CStr(IdCta) + "," + Replace(CStr(Var), ",", ".") + ",#" + _
        stFechaSQL(Fecha) + "#, '" + Detalle + "')"
End Sub

'acepta valores negativos para no tener que hacer
'if para ver si se paga o se cobra por ejemplo
Public Function Asiento(DebitosSepPorBarras As String, MontosD As String, _
    CreditosSepPorBarras As String, MontosC As String, _
    Optional Libro As String = "LibroDiario", _
    Optional Detalle As String = "", Optional IdAs As Long = 0) As Long
        '0 JOIIIIIIA
        '1 si alguno de los argumentos es vacio
        '2 si no coinciden las cantidades de montos y de cuentas
        '3 si debitos y creditos no son iguales
        
    'Si IdAs es 0 es por que quiero AGREGAR un asiento
    '       si es positivo es para CAMBIAR un el asiento IdAs por uno nuevo con ese nro
    '       si es negativo es para AÑADIRLE movimientos al asiento -IdAs existente
    
    Dim IdAsiento As Long
    Dim Debitos() As String, MDebitos() As String
    Dim Creditos() As String, MCreditos() As String
    Dim Fech As Date
    
    If IdAs = 0 Then
        IdAsiento = GetUltIDAsientoMasUno(Libro)
        Fech = Date
    Else
        IdAsiento = Abs(IdAs)
        If Not IsDate(GetValInRS(Libro, "Fecha", "IdAsiento = " + CStr(IdAsiento))) Then
            Fech = Date
        Else
            Fech = CDate(GetValInRS(Libro, "Fecha", "IdAsiento = " + CStr(IdAsiento)))
        End If
        'si IDAs es positivo borro aca nomas el asiento original
        If IdAs > 0 Then
            mCN.Execute "DELETE * FROM " + Libro + " WHERE IdAsiento = " + CStr(IdAs)
        End If
    End If
    
    If DebitosSepPorBarras = "" Or CreditosSepPorBarras = "" _
        Or MontosD = "" Or MontosC = "" Then
        
        Asiento = 1 'algo se cargo mal
        Exit Function
    End If
    
    Debitos = Split(DebitosSepPorBarras, "/")
    Creditos = Split(CreditosSepPorBarras, "/")
    MDebitos = Split(MontosD, "/")
    MCreditos = Split(MontosC, "/")
    
    If UBound(Debitos) <> UBound(MDebitos) Or UBound(Creditos) <> UBound(MCreditos) Then
        Asiento = 2 'falta algun monto para alguna cuenta
        Exit Function
    End If
    
    'controlo que montos de debitos y creditos sean iguales
    If ControlDyC(MDebitos, MCreditos) = False Then
        Asiento = 3 'debitos y creditos no son iguales
        Exit Function
    End If
    
        'registro debitos
    For I = 0 To UBound(Debitos)
        If Abs(CSng(MDebitos(I))) > 0.001 Then
            AgregarMov IdAsiento, CLng(Debitos(I)), CSng(MDebitos(I)), Fech, _
                Detalle, Libro
        End If
    Next
        
        'registo creditos
    For I = 0 To UBound(Creditos)
        If Abs(CSng(MCreditos(I))) > 0.001 Then
            AgregarMov IdAsiento, CLng(Creditos(I)), -CSng(MCreditos(I)), _
                Fech, Detalle, Libro
        End If
    Next I
    
End Function

'saca los codigos de los 3 niveles separados por barras por ej 013/024/000
Private Function ControlDyC(Debitos() As String, Creditos() As String) As Boolean
    Dim Ii As Long, jj As Long, Deb As Single, Cred As Single
    Deb = 0
    Cred = 0
    
    ControlDyC = False
    
    For Ii = 0 To UBound(Debitos)
        Deb = Deb + CSng(Debitos(Ii))
    Next
    For jj = 0 To UBound(Creditos)
        Cred = Cred + CSng(Creditos(jj))
    Next

    'puede haber una diferencia como maximo de 50 cents
    If Abs(Deb - Cred) < 0.5 Then ControlDyC = True
    
End Function

Public Sub PasarSubDiario()
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'resumo todo lo que haya en subdiario y lo pasa a librodiario) en un solo asiento
    
    Dim SS As String, IdAsiento As Long
    Dim RsR As New ADODB.Recordset
    
    IdAsiento = GetUltIDAsientoMasUno
    SS = "SELECT IdCuenta, Sum(Variacion) AS " + _
        "SumaDeVariacion From LibroSubDiario GROUP BY IdCuenta"
    
    RsR.Open SS, mCN, adOpenStatic, adLockOptimistic
    
    If RsR.RecordCount = 0 Then
        RsR.Close
        Set RsR = Nothing
        Exit Sub
    End If
    
    RsR.MoveFirst
    Do While Not RsR.EOF
        mCN.Execute "INSERT INTO LibroDiario (IdAsiento,IdCuenta,Variacion,Fecha,Detalle) " + _
            "VALUES (" + CStr(IdAsiento) + ", " + CStr(RsR("idcuenta")) + "," + _
            Replace(CStr(RsR("SumaDeVariacion")), ",", ".") + _
            ",#" + stFechaSQL(Date) + "#,'" + _
            "Subdiario Compras-Ventas al " + FormatDateTime(Date, vbShortDate) + "')"
            
        RsR.MoveNext
    Loop
    
    RsR.Close
    Set RsR = Nothing
    
    'ahora borro todos los asientos del subdiario
    mCN.Execute "DELETE FROM LibroSubDiario"
    
End Sub

Public Function ExisteNCuenta(NombreCuenta As String) As Long
    '0 si no hay ninguna cuenta con ese nombre
    'otro nro si si
    Dim RSBusca As New ADODB.Recordset
    Dim Res As Long
    
    RSBusca.Open "SELECT Nombre FROM tblCuentas WHERE nombre = '" + _
        NombreCuenta + "'", mCN, adOpenStatic, adLockReadOnly
    ExisteNCuenta = RSBusca.RecordCount
    RSBusca.Close
    Set RSBusca = Nothing
    
End Function

Public Function FechaUltimoCierre() As Date
    Dim IDAsCi As Long, Fecha As Date
    
    IDAsCi = GetUltIdCierreCaja
    
    If IDAsCi = 0 Then 'por ahora lo dejo asi
        Fecha = #1/1/1901# 'para que no haga lo del capital inicial
    Else
        Fecha = CDate(GetValInRS("LibroDiario", "Fecha", "IdAsiento = " + CStr(IDAsCi), True))
    End If
    
    FechaUltimoCierre = Fecha

End Function

Public Function TipoCuenta(IdCta As Long) As Long
    'busca el id padre
    '1 activo, 2 pasivo, 3 perdida,4 ganancia, 5 PN
    Dim RSq As New ADODB.Recordset
    Dim tmP As Long 'graba el idant
    
    If IdCta < 5 Then
        TipoCuenta = IdCta
        Set RSq = Nothing
        Exit Function
    End If
    
    RSq.Open "SELECT Id,IdAnt FROM tblCuentas WHERE Id = " + CStr(IdCta), _
        mCN, adOpenStatic, adLockReadOnly
    
    If RSq.RecordCount = 0 Then
        TipoCuenta = IdCta
    Else
        If RSq("IdAnt") = 0 Then
            TipoCuenta = IdCta
        Else
            TipoCuenta = TipoCuenta(RSq("IdAnt"))
        End If
    End If
    
    RSq.Close
    Set RSq = Nothing
End Function

've la variacion en comparacion con el ultimo cierre de caja con el que tengo ahora
Public Function MovCerrados(idCuenta As Long) As Single
    
    MovCerrados = GetValInRS("cstBalanceCerrados", "SumadeVariacion", _
        "ID = " + CStr(idCuenta), False)

End Function

've la variacion en comparacion con el ultimo cierre de caja con el que tengo ahora
Public Function MovCerradosSub(idCuenta As Long) As Single
    
    UltVarCuentaSub = GetValInRS("cstBalanceSubCerrados", "SumadeVariacion", _
        "ID = " + CStr(idCuenta), False)
End Function

Public Function GetSaldo(idCuenta As Long, Optional Libro As String = "Todos") As Single
    Dim Diar As Single, SubDiar As Single
    
    Diar = GetValInRS("cstBalance", "SumadeVariacion", _
        "ID = " + CStr(idCuenta), False)
    SubDiar = GetValInRS("cstBalanceSub", "SumadeVariacion", _
        "ID = " + CStr(idCuenta), False)
        
    Select Case Libro
        Case "Todos"
            GetSaldo = Diar + SubDiar
        Case "LibroDiario"
            GetSaldo = Diar
        Case "LibroSubDiario"
            GetSaldo = SubDiar
        Case Else
            GetSaldo = Diar + SubDiar
    End Select
    
End Function

Public Function GetSaldoHasta(idCuenta As Long, Hasta As Date, _
    Optional Libro As String = "Todos", Optional Inclusive As Boolean = False) As Single
    'da el saldo hasta tal fecha SIN incluir esa fecha o sea < que esa fecha
    Dim Diar As Single, SubDiar As Single
    Dim Carac As String
    
    Carac = " <"
    If Inclusive Then Carac = Carac + "="
    
    Diar = SumarValInRS("LibroDiario", "Variacion", "IdCuenta = " + CStr(idCuenta) + _
        " AND Fecha" + Carac + " #" + stFechaSQL(Hasta) + " 00:00:00#")
    SubDiar = SumarValInRS("LibroSubDiario", "Variacion", "IdCuenta = " + CStr(idCuenta) + _
        " AND Fecha" + Carac + " #" + stFechaSQL(Hasta) + " 00:00:00#")
        
    Select Case Libro
        Case "Todos"
            GetSaldoHasta = Diar + SubDiar
        Case "LibroDiario"
            GetSaldoHasta = Diar
        Case "LibroSubDiario"
            GetSaldoHasta = SubDiar
        Case Else
            GetSaldoHasta = Diar + SubDiar
    End Select

End Function

Public Function ABSValor(IdCta As Long, Valor As Single) As Single
    Dim TipoC As Long, Suma As Single
    
    TipoC = TipoCuenta(IdCta)
    Suma = Valor
    
    Select Case TipoC
        Case 1, 3 'activo y perdida deben tener saldos deudores (positivo en librodiario)
            ABSValor = Suma 'no hago nada mantengo su signo
        Case 2, 4, 5
            ABSValor = -Suma 'cambio el signo de los saldos
    End Select
End Function

Public Function GetSaldoHastaPorNroAsiento(idCuenta As Long, Hasta As Long, _
    Optional Libro As String = "Todos") As Single
    'da el saldo hasta tal fecha SIN incluir esa fecha o sea < que esa fecha
    Dim Diar As Single, SubDiar As Single
    Dim Carac As String
    
    Carac = " <"
    If Inclusive Then Carac = Carac + "="
    
    Diar = SumarValInRS("LibroDiario", "Variacion", "IdCuenta = " + CStr(idCuenta) + _
        " AND IdAsiento" + Carac + CStr(Hasta))
    SubDiar = SumarValInRS("LibroSubDiario", "Variacion", "IdCuenta = " + CStr(idCuenta) + _
        " AND IdAsiento" + Carac + CStr(Hasta))
        
    Select Case Libro
        Case "Todos"
            GetSaldoHastaPorNroAsiento = Diar + SubDiar
        Case "LibroDiario"
            GetSaldoHastaPorNroAsiento = Diar
        Case "LibroSubDiario"
            GetSaldoHastaPorNroAsiento = SubDiar
        Case Else
            GetSaldoHastaPorNroAsiento = Diar + SubDiar
    End Select

End Function

Public Function SumarConSubcuentas(ByVal IdCta As Long, _
    Optional ConCerrados As Boolean = True) As Single
    
    Dim Hijos() As String, Suma As Single
    
    Suma = 0
    
    'sumo el padre y los hijos los hago que se sumen con esta misma funcion
    Suma = GetSaldo(IdCta)
    If ConCerrados = False Then 'resto los ultimos movimientos
        Suma = Suma - MovCerrados(IdCta)
        Suma = Suma - MovCerradosSub(IdCta)
    End If
    
    'ahora busco los hijos que tengan en idant el id de la cuenta
    Hijos = GetCuentas(IdCta)
    Dim jj As Long
    
    If UBound(Hijos) > 0 Then
        For jj = 1 To UBound(Hijos)
            Suma = Suma + SumarConSubcuentas(Hijos(jj), ConCerrados)
        Next
    End If
    
    SumarConSubcuentas = Suma

End Function

Public Function ABSSumarconSubcuentas(IdCta As Long, _
    Optional ConCerrados As Boolean = True) As Single
    
    Dim TipoC As Long, Suma As Single
    
    TipoC = TipoCuenta(IdCta)
    Suma = SumarConSubcuentas(IdCta, ConCerrados)
    
    Select Case TipoC
        Case 1, 3 'activo y perdida deben tener saldos deudores (positivo en librodiario)
            ABSSumarconSubcuentas = Suma 'no hago nada mantengo su signo
        Case 2, 4, 5
            ABSSumarconSubcuentas = -Suma 'cambio el signo de los saldos
    End Select

End Function

Public Function UltVarCuenta(idCuenta As Long) As Single
    UltVarCuenta = GetSaldo(idCuenta) - MovCerrados(idCuenta)
End Function

Public Function AsientoRNA() As String()
    'devuelve lo que deberia ser el asiento de cierre de resultados
    'en forma de matriz de esta manera (1)IDcuenta | debe (2) IDcuenta | -haber ,etc
    'de manera que el programa le pregunte primero si desea hacer el asiento
    'y mostrarselo para que decida y ahi si usar el sub CerrarResultados
    'debe estar previamente pasados los mov del subdiario!!!!!!!!!
    
    Dim RSRna As New ADODB.Recordset
    Dim Asiento() As String, H As Long, RNA As Single, TypC As Long
    Dim Hh As String
        
    Hh = "SELECT LibroDiario.IDCuenta, Sum(LibroDiario.Variacion) AS SumVar " + _
        "FROM tblCuentas INNER JOIN LibroDiario ON tblCuentas.ID = " + _
        "LibroDiario.IDCuenta GROUP BY LibroDiario.IDCuenta"
'        "WHERE CerroCaja <> 'Resultados'" + _
'        "GROUP BY LibroDiario.IDCuenta"

        
    RNA = 0: H = 0
    ReDim Preserve Asiento(H)
    Asiento(H) = "Nada"
    
    RSRna.Open Hh, mCN, adOpenStatic, adLockReadOnly
    
    If RSRna.RecordCount > 0 Then
        RSRna.MoveFirst
        
        Do While Not RSRna.EOF
            TypC = TipoCuenta(RSRna("IdCuenta"))
            If TypC = 3 Or TypC = 4 Then 'si es de resultados debe ser anulado
                If Abs(NoNuloN(RSRna("SumVar"))) >= 0.1 Then
                    RNA = RNA + NoNuloN(RSRna("SumVar"))
                    H = H + 1
                    ReDim Preserve Asiento(H)
                    Asiento(H) = CStr(RSRna("IDCuenta")) + " | " + _
                        CStr(-NoNuloN(RSRna("SumVar")))
                End If
            End If
            RSRna.MoveNext
        Loop
        
        'falta nomas la cuenta RNA (16 es una cuenta intocable)
        If UBound(Asiento) > 0 Then 'si no hay nada no pongo nada
            H = H + 1
            ReDim Preserve Asiento(H)
            Asiento(H) = "16 | " + CStr(RNA)
        End If
    End If
    
    RSRna.Close
    Set RSRna = Nothing
    
    AsientoRNA = Asiento
End Function


Public Sub CerrarResultados()
    'Unicamente hago que se haga un asiento de cierre de las cuentas de resultado
    'y lo tire a RNA y veamos que pasa
    Dim Asiento() As String, NroAsiento As Long
    Dim sP() As String, I As Long
    
    Asiento = AsientoRNA 'debo estar 100% seguro que sumadebitos=sumacreditos
                         'estoy un 99% por eso sigo nomas
    NroAsiento = GetUltIDAsientoMasUno
    
    For I = 1 To UBound(Asiento)
        sP = Split(Asiento(I), " | ")
        
        AgregarMov NroAsiento, CLng(sP(0)), CSng(sP(1)), Date, _
            "Cierre Resultados al " + FormatDateTime(Date, vbShortDate)
    Next I
    
    'bueno quedo todo listo solo falta marcar como Resultados en CerroCaja
    mCN.Execute "UPDATE LibroDiario SET CerroCaja = 'Resultados' " + _
        "WHERE IDAsiento <= " + CStr(NroAsiento)
End Sub

Public Function ListarAsientos(IdAsientoDe As Long, IdAsientoA As Long, LvW As Object, _
    Optional Libro As String = "LibroDiario") As Date
    'object tiene 4 columnas: IdCuenta,Nombre, Debe, Haber.
    'Fecha da como resultado, si no existe el asiento -> fecha = 1/1/1900
    
    Dim S As String, tmP As Long, Resp As Date
    Dim RSq As New ADODB.Recordset
    
    S = "SELECT " + Libro + ".Fecha, " + Libro + ".IDCuenta, " + _
        "tblCuentas.Nombre , " + Libro + ".Detalle, " + Libro + ".Variacion " + _
        "FROM tblCuentas INNER JOIN " + Libro + " ON " + _
        "tblCuentas.ID = " + Libro + ".IDCuenta " + _
        "WHERE " + Libro + ".IdAsiento BETWEEN " + CStr(IdAsientoDe) + _
        " AND " + CStr(IdAsientoA) + _
        " GROUP BY " + Libro + ".Fecha, " + Libro + ".IDCuenta, " + _
        "tblCuentas.Nombre, " + Libro + ".Detalle, " + Libro + ".Variacion"
    
    RSq.Open S, mCN, adOpenStatic, adLockReadOnly
    
    tmP = 0
    LvW.ListItems.Clear
    
    If RSq.RecordCount > 0 Then
        RSq.MoveFirst
        Do While Not RSq.EOF
            Resp = CDate(RSq("Fecha"))
            
            tmP = tmP + 1
            LvW.ListItems.Add tmP
            
            LvW.ListItems(tmP).Text = CStr(RSq("IdCuenta"))
            LvW.ListItems(tmP).SubItems(1) = RSq("Nombre")
            
            If CSng(RSq("Variacion")) < 0 Then
                LvW.ListItems(tmP).SubItems(3) = FormatCurrency(-CSng(RSq("Variacion")), , , , vbFalse)
            Else
                LvW.ListItems(tmP).SubItems(2) = FormatCurrency(CSng(RSq("Variacion")), , , , vbFalse)
            End If
            
            RSq.MoveNext
        Loop
    Else
        Resp = #1/1/1900#
    End If
        
    RSq.Close
    Set RSq = Nothing
    
    ListarAsientos = Resp
End Function

Public Function GetUltIdCierreResultados() As Long
    GetUltIdCierreResultados = GetTop1Rs("LibroDiario", "IdAsiento", , _
        "Detalle LIKE '%Cierre Resultados%'", False)
End Function

Public Function GetUltIdCierreCaja() As Long
    GetUltIdCierreCaja = GetTop1Rs("LibroDiario", "IDAsiento", , _
        "Detalle LIKE '%Cierre de Caja%'", False)
        
End Function

Public Function ListarMovCuentaPorNroAsiento(idCuenta As Long, Desde As Long, _
    Hasta As Long, _
    LvW As Object, Optional Libro As String = "LibroDiario", _
    Optional ArrastraSaldo As Boolean = True) As Single
    'object tiene 4 columnas: Fecha, IdAsiento,Detalle y variacion
    'da como resultado el saldo
    
    Dim Arrastre As Single, tmP As Long
    Dim RsL As New ADODB.Recordset
        
    If Hasta = 0 Then Hasta = GetUltIDAsientoMasUno(Libro)
    
    Arrastre = 0: tmP = 0
    LvW.ListItems.Clear
    
    If ArrastraSaldo Then
        Arrastre = GetSaldoHastaPorNroAsiento(idCuenta, Desde, Libro)
        tmP = LvW.ListItems.Count + 1
        LvW.ListItems.Add tmP
        
        LvW.ListItems(tmP).Text = FormatDateTime(Date, vbShortDate)
        LvW.ListItems(tmP).SubItems(1) = ""
        LvW.ListItems(tmP).SubItems(2) = "Saldo hasta asiento nro " + CStr(Desde)
        LvW.ListItems(tmP).SubItems(3) = FormatCurrency(Arrastre, , , , vbFalse)
    End If
    
    'ahora si vamos con los mov
    RsL.Open "SELECT * FROM " + Libro + " WHERE IdCuenta = " + CStr(idCuenta) + _
        " AND IDAsiento BETWEEN " + CStr(Desde) + _
        " AND " + CStr(Hasta) + " ORDER BY FECHA DESC", _
        mCN, adOpenStatic, adLockReadOnly
        
    If RsL.RecordCount > 0 Then
        RsL.MoveFirst
        
        Do While Not RsL.EOF
            If EsCero(NoNuloN(RsL("Variacion"))) = False Then
                tmP = tmP + 1
                LvW.ListItems.Add tmP
            
                LvW.ListItems(tmP).Text = FormatDateTime(RsL("Fecha"), vbShortDate)
                LvW.ListItems(tmP).SubItems(1) = CStr(RsL("IdAsiento"))
                LvW.ListItems(tmP).SubItems(2) = NoNuloS(RsL("Detalle"))
                LvW.ListItems(tmP).SubItems(3) = FormatCurrency(NoNuloS(RsL("Variacion")), , , , vbFalse)
                
                Arrastre = Arrastre + NoNuloN(RsL("Variacion"))
            End If
            RsL.MoveNext
        Loop
    End If
    
    RsL.Close
    Set RsL = Nothing
    
    'este el saldo hasta
    ListarMovCuentaPorNroAsiento = Arrastre
End Function


Public Function ListarMovCuenta(idCuenta As Long, Desde As Date, Hasta As Date, _
    LvW As Object, Optional Libro As String = "LibroDiario", _
    Optional ArrastraSaldo As Boolean = True, Optional ValorAbs As Boolean) As Single
    'object tiene 4 columnas: Fecha, IdAsiento,Detalle y variacion
    'da como resultado el saldo
    
    Dim Arrastre As Single, tmP As Long
    Dim RsL As New ADODB.Recordset
    
    If Hasta = 0 Then Hasta = Date
        
    Arrastre = 0: tmP = 0
    LvW.ListItems.Clear
    
    If ArrastraSaldo Then
        Arrastre = GetSaldoHasta(idCuenta, Desde, Libro)
        tmP = LvW.ListItems.Count + 1
        LvW.ListItems.Add tmP
        
        LvW.ListItems(tmP).Text = FormatDateTime(Desde, vbShortDate)
        LvW.ListItems(tmP).SubItems(1) = ""
        LvW.ListItems(tmP).SubItems(2) = "Saldo al " + FormatDateTime(Desde, vbShortDate)
        LvW.ListItems(tmP).SubItems(3) = FormatCurrency(Arrastre, , , , vbFalse)
    End If
    
    'ahora si vamos con los mov
    RsL.Open "SELECT * FROM " + Libro + " WHERE IdCuenta = " + CStr(idCuenta) + _
        " AND Fecha BETWEEN #" + stFechaSQL(Desde) + " 00:00:00# " + _
        "AND #" + stFechaSQL(Hasta) + " 23:59:59# ORDER BY FECHA ASC, " + _
        "IdAsiento ASC", _
        mCN, adOpenStatic, adLockReadOnly
        
    If RsL.RecordCount > 0 Then
        RsL.MoveFirst
        
        Do While Not RsL.EOF
            If EsCero(NoNuloN(RsL("Variacion"))) = False Then
            
                tmP = tmP + 1
                LvW.ListItems.Add tmP
            
                LvW.ListItems(tmP).Text = FormatDateTime(RsL("Fecha"), vbShortDate)
                LvW.ListItems(tmP).SubItems(1) = CStr(RsL("IdAsiento"))
                LvW.ListItems(tmP).SubItems(2) = NoNuloS(RsL("Detalle"))
                If ValorAbs Then
                    LvW.ListItems(tmP).SubItems(3) = FormatCurrency( _
                        ABSValor(idCuenta, NoNuloN(RsL("Variacion"))), , , , vbFalse)
                Else
                    LvW.ListItems(tmP).SubItems(3) = FormatCurrency(NoNuloN(RsL("Variacion")), , , , vbFalse)
                End If
                Arrastre = Arrastre + NoNuloN(RsL("Variacion"))
            End If
            
            RsL.MoveNext
        Loop
    End If
    
    RsL.Close
    Set RsL = Nothing
    
    'este el saldo hasta
    ListarMovCuenta = Arrastre
End Function

Public Function Distribuir(IdCierre As Long, Socios As Single, Capitalizar As Single)
    'registro la distribucion
    mCN.Execute "INSERT INTO Distribuciones (IdCierre,Socios,Capitalizar) VALUES (" + _
        CStr(IdCierre) + "," + Replace(CStr(Socios), ",", ".") + "," + _
        Replace(CStr(Capitalizar), ",", ".") + ")"

End Function

Public Function VerDistribucion(IdCierre As Long, Cuenta As String) As Single
    'cuenta es socios o capitalizar, por ahora
    
    VerDistribucion = GetValInRS("Distribuciones", Cuenta, _
        "IdCierre = " + CStr(IdCierre), False)

End Function

Public Function CerrarCaja() As Single
    '(1) pasa todo del subdiario
    '(2) marca todo el diario en CerroCaja -> SI
    '(3) Da como resultado el valor del caja al cierre

    '(1)
    PasarSubDiario
    
    '(2)
    mCN.Execute "UPDATE LibroDiario SET CerroCaja = 'Si' WHERE CerroCaja = 'No'"
    
    '(3)
    CerrarCaja = GetSaldo(78)
    
End Function

Public Function CerroCaja() As Boolean
    'primero obligo que cierre de caja antes de hacer resultados
    Dim CerroC As Boolean
    
    CerroC = False 'predeterminado
    
    If ContarReg("SELECT * FROM LibroSubDiario") = 0 Then
        If ContarReg("SELECT * FROM LibroDiario WHERE CerroCaja = 'No'") = 0 Then
            CerroC = True
        End If
    
    End If
    
    CerroCaja = CerroC
End Function

Public Sub CierresViejos()
    frmCierresViejos.AbrirDatos mArchMDB, mPsW
End Sub

Public Sub ListarEgresos(LvW As Object, Top As Long, Optional Desde As Date = #1/1/1980#)
    Dim RSeg As New ADODB.Recordset
    Dim Ttop As Long, tmP As Long, Reng As Long
    
    'object tiene 5 columnas fecha,IDAs, NbCuenta,Detalle, variacion
    'le da bola a la fecha si pone top=0
    'va a ir restando a medida que vaya agregando
    
    LvW.ListItems.Clear
    
    RSeg.Open "SELECT LibroDiario.Fecha, LibroDiario.IDCuenta, tblCuentas.Nombre, " + _
        "LibroDiario.Variacion, LibroDiario.IDAsiento, LibroDiario.Detalle " + _
        "FROM tblCuentas INNER JOIN LibroDiario ON tblCuentas.ID = " + _
        "LibroDiario.IDCuenta WHERE Fecha >= #" + stFechaSQL(Desde) + "# AND " + _
        "Detalle NOT LIKE '%Cierre Resultados%' " + _
        "ORDER BY LibroDiario.Fecha DESC", mCN, adOpenStatic, adLockReadOnly
     
    If Top = 0 Then
        Ttop = RSeg.RecordCount
    Else
        Ttop = Top
    End If
    Reng = 0
        
    If RSeg.RecordCount > 0 Then
        RSeg.MoveFirst
        Do While Not (RSeg.EOF Or Ttop = 0)
            tmP = TipoCuenta(RSeg("IdCuenta"))
            If tmP = 3 Or tmP = 4 Then
                If EsCaja(RSeg("idCuenta")) = 1 Then
                        'nada
                Else
                    ' los de sueldos no!
                    'ya se borran en initialize
                    Reng = Reng + 1
                    
                    LvW.ListItems.Add Reng
                    
                    LvW.ListItems(Reng).Text = FormatDateTime(RSeg("Fecha"), vbShortDate)
                    LvW.ListItems(Reng).SubItems(1) = RSeg("IDAsiento")
                    LvW.ListItems(Reng).SubItems(2) = RSeg("Nombre")
                    LvW.ListItems(Reng).SubItems(3) = RSeg("Detalle")
                    LvW.ListItems(Reng).SubItems(4) = FormatCurrency(RSeg("Variacion"), , , , vbFalse)
                    
                    Ttop = Ttop - 1
                
                End If
            End If
            
            RSeg.MoveNext
        Loop
    End If
    
    RSeg.Close
    Set RSeg = Nothing

End Sub

Private Sub ControlarSueldos()
    'pasa todos como EsCaja = 1 para que no se muestre en egresos
    Dim I As Long, RSeg As New ADODB.Recordset, tmP As Long
    
    RSeg.Open "SELECT * FROM tblCuentas", mCN, adOpenStatic, adLockReadOnly
    
    RSeg.MoveFirst
    Do While Not (RSeg.EOF Or Ttop = 0)
        tmP = TipoCuenta(RSeg("IdCuenta"))
        If tmP = 3 Then
            'si tiene esto es que es Sueldo (supongo XXX)
            If Left(tmP, 7) = "Sueldo " Then
                mCN.Execute "UPDATE tblCuentas SET EsCaja = 1 WHERE ID = " + _
                    RSeg("ID")
            End If
            RSeg.MoveNext
        End If
    Loop
    
    RSeg.Close
    Set RSeg = Nothing
End Sub

Private Function EsCaja(idCuenta As Long) As Long
    'para cuentas de resultados, devuelve 1 si no genera caja, 0 si si
    Dim tmP As Long
    Dim rS As New ADODB.Recordset
    
    tmP = 0 ' predeterminado
    
    rS.Open "SELECT Caja FROM tblCuentas WHERE ID = " + _
        CStr(idCuenta), mCN, adOpenStatic, adLockReadOnly
    
    If rS.RecordCount > 0 Then
        tmP = rS("Caja")
    End If
    
    EsCaja = tmP
    
    rS.Close
    Set rS = Nothing
End Function

Public Sub LimpiarCierres(DiasAtras As Long)
    mCN.Execute "DELETE FROM CierresViejos WHERE Fecha <#" + _
        stFechaSQL(Date - DiasAtras) + "#"
End Sub

Public Function ResumirAsientos(DiasAtras As Long, Optional Libro As String = _
    "LibroDiario") As Long
        'mismas respuestas que Asiento
    Dim Resp As Long
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'resumo todo lo que haya en tblmovimientos (librodiario) a 1 solo asiento
    ' los asientos viejos hasta lo pida el usuario
    Dim DiaD As Date, SS As String
    Dim RsR As New ADODB.Recordset
    
    Dim strDeb As String, strCred As String
    Dim strMonD As String, strMonH As String
    
    strDeb = "": strCred = ""
    strMonD = "": strMonH = ""
    
    DiaD = Date - DiasAtras
    
    SS = "SELECT " + Libro + ".IDCuenta, Sum(" + Libro + ".Variacion) AS Var " + _
        "FROM " + Libro + " " + _
        "WHERE " + Libro + ".Fecha < #" + stFechaSQL(DiaD) + " 00:00#" + _
        "GROUP BY " + Libro + ".IDCuenta"

    RsR.Open SS, mCN, adOpenStatic, adLockOptimistic
    
    If RsR.RecordCount = 0 Then
        RsR.Close
        Set RsR = Nothing
        Exit Function
    End If
    
    RsR.MoveFirst
    Do While Not RsR.EOF
        If EsCero(RsR("Var")) = False Then
            If CSng(RsR("Var")) < 0 Then 'es negativo -> cuenta acreedora
                If strCred <> "" Then
                    strCred = strCred + "/"
                    strMonH = strMonH + "/"
                End If
                strCred = strCred + CStr(RsR("IdCuenta"))
                strMonH = strMonH + CStr(-CSng(RsR("Var")))
            Else 'es positivo -> Cuenta Deudora
                If strDeb <> "" Then
                    strDeb = strDeb + "/"
                    strMonD = strMonD + "/"
                End If
                strDeb = strDeb + CStr(RsR("IdCuenta"))
                strMonD = strMonD + CStr(RsR("Var"))
            End If
        End If
        RsR.MoveNext
    Loop
    
    RsR.Close
    Set RsR = Nothing
    
    Resp = Asiento(strDeb, strMonD, strCred, strMonH, , "Asiento Resumen Libro " + _
        Right(Libro, Len(Libro) - 5) + " por asientos anteriores al " + CStr(DiaD))
    'ahora borro todos los asientos individuales entre esas fechas
    If Resp = 0 Then 'todo ok
        mCN.Execute "DELETE FROM " + Libro + " WHERE Fecha < #" + _
            stFechaSQL(DiaD) + " 00:00#"
    End If
    
    ResumirAsientos = Resp
End Function

'!!!!!!!!!!!!!!!!!!!!!!GLOBALES DE LA DLL !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!!!!!!!!!!!!!!!!!!!GLOBALES DE LA DLL !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!!!!!!!!!!!!!!!!!!!GLOBALES DE LA DLL !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!!!!!!!!!!!!!!!!!!!GLOBALES DE LA DLL !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Private Function EsCero(Nro As Single) As Boolean
    If Abs(Nro) < 0.01 Then
        EsCero = True
    Else
        EsCero = False
    End If
End Function

Public Function GetValInRS(Tabla As String, Campo As String, sWhere As String, _
    Optional EsString As Boolean = True)
    'devuelve el valor de un campo en una tabla en  un campo especificado dada
    'determinada condicion (supongo que es solo 1)
    
    Dim rS As New ADODB.Recordset
    rS.Open "Select " + Campo + " FROM " + Tabla + " WHERE " + sWhere, _
        mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        If EsString Then
            GetValInRS = ""
        Else
            GetValInRS = 0
        End If
    Else
        If EsString Then
            GetValInRS = NoNuloS(rS.Fields(Campo))
        Else
            GetValInRS = NoNuloN(rS.Fields(Campo))
        End If
    End If
    
    rS.Close
    Set rS = Nothing
End Function

Private Function SumarValInRS(Tabla As String, Campo As String, _
    sWhere As String) As Single
    
    'devuelve la suma de un campo en una tabla en  un campo especificado dada
    'determinada condicion
    Dim Suma As Single
    Suma = 0
    
    Dim rS As New ADODB.Recordset
    rS.Open "Select " + Campo + " FROM " + Tabla + " WHERE " + sWhere, _
        mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        
    Else
        rS.MoveFirst
        Do While Not rS.EOF
            Suma = Suma + NoNuloN(rS.Fields(Campo))
            
            rS.MoveNext
        Loop
        
    End If
    
    SumarValInRS = Suma
    rS.Close
    Set rS = Nothing
End Function

Public Function GetTop1Rs(Tabla As String, Campo As String, _
    Optional ORDERBY As String = "DESC", Optional WHERE As String = "NINGUNO", _
    Optional EsString As Boolean = False)
        
    'devuelve el top 1 de un campo determinado
    Dim rS As New ADODB.Recordset
    Dim tmpW As String
    
    If WHERE = "NINGUNO" Then
        tmpW = ""
    Else
        tmpW = " WHERE " + WHERE
    End If
    
    rS.CursorLocation = adUseClient
    rS.Open "Select TOP 1 " + Campo + _
        " FROM " + Tabla + tmpW + _
        " ORDER BY " + Campo + " " + ORDERBY, _
        mCN, adOpenStatic, adLockReadOnly
        
    If rS.RecordCount = 0 Then
        If EsString Then
            GetTop1Rs = ""
        Else
            GetTop1Rs = 0
        End If
        
        rS.Close
        Set rS = Nothing
        Exit Function
    End If
    
    If EsString Then
        GetTop1Rs = NoNuloS(rS.Fields(Campo))
    Else
        GetTop1Rs = NoNuloN(rS.Fields(Campo))
    End If
    
    rS.Close
    Set rS = Nothing
End Function

Public Function NoNuloS(S) As String
    If IsNull(S) Then
        NoNuloS = ""
    Else
        NoNuloS = S
    End If
End Function

Public Function NoNuloN(J) As Single
    If IsNumeric(J) Then
        NoNuloN = J
    Else
        NoNuloN = 0
    End If
End Function

Public Function stFechaSQL(Fecha As Date) As String
    Dim H() As String
    H = Split(CStr(Fecha), "/")
    stFechaSQL = H(1) + "/" + H(0) + "/" + H(2)
    
End Function

Public Sub CargarCombo(CMB As Object, sqlText As String, _
    CamposSeparadosPorComas As String, Optional Separador As String = "\")
    'ZZZZ pasar al completo
    
    'CamposSeparadosPorComas es una lista separada por comas de los campos. _
        Ademas se le puede agregar al final _
        /n al final para indicar que es numero _
        /f para fechas _
        /$ para currency _
        predeterminado es string
    
    Dim Campos() As String
    Campos = Split(CamposSeparadosPorComas, ",")
        
    Dim rS As New ADODB.Recordset
    rS.Open sqlText, mCN, adOpenStatic, adLockReadOnly
    Dim S As String, AA As Long
    CMB.Clear
    If rS.RecordCount = 0 Then Exit Sub
    rS.MoveFirst
    Do While Not rS.EOF
        S = ""
        For AA = 0 To UBound(Campos)
            Dim Ult2 As String 'ultimos dos caracteres del campo
            Ult2 = Right(Campos(AA), 2)
            Dim NombreRealCampo As String
            Select Case Ult2
                Case "/n"
                    NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                    S = S + CStr(NoNuloN(rS(NombreRealCampo)))
                Case "/f"
                    NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                    S = S + CStr(rS(NombreRealCampo))
                Case "/$"
                    NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                    S = S + FormatCurrency(rS(NombreRealCampo), , , , vbFalse)
                Case Else
                    S = S + NoNuloS(rS(Campos(AA)))
            End Select
            'si no es el ultimo poner la barra separadora
            If AA < UBound(Campos) Then S = S + Separador
        Next AA
        CMB.AddItem S
        rS.MoveNext
    Loop
    rS.Close
    Set rS = Nothing
    CMB.ListIndex = 0
End Sub

Private Function ContarReg(sSQL As String) As Long
    'cuenta la cantidad de registros de un RS
    'sirve para duplicados por ejemplo
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    
    rS.Open sSQL, mCN, adOpenStatic, adLockReadOnly
    ContarReg = rS.RecordCount
    
    rS.Close
    Set rS = Nothing

End Function





VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTbrAccesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'modulo para registrar:
'Private CN As New ADODB.Connection
Public mDBFilename As String 'ubicacion de la base

'EVENTOS NECESARIOS (no cambiar los nros)
'1 - Asignar Eventos
'5 - Abrir sistema
'6 - Cerrar Sesion

Public Sub Conectar()
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _
        mDBFilename + "; Jet OLEDB:Database Password=zuliani"
    CN.Open
End Sub

Public Sub Desconectar()
    If CN.State = adStateOpen Then CN.Close
End Sub

Public Property Let DBFilename2(NewDB2 As String)
    mDBFilename = NewDB2
End Property

Public Property Get DBFilename2() As String
    DBFilename2 = mDBFilename
End Property

Private Function ExisteNCuenta(Campo As String, Tabla As String, _
    NombreCuenta As String) As Long
    '0 si no hay ninguna cuenta con ese nombre
    'otro nro si si
    Dim RSBusca As New ADODB.Recordset
    Dim Res As Long
    
    RSBusca.Open "SELECT " + Campo + " FROM " + Tabla + " WHERE " + _
        Campo + " = '" + _
        NombreCuenta + "'", CN, adOpenStatic, adLockReadOnly
    
    ExisteNCuenta = RSBusca.RecordCount
    
    RSBusca.Close
    Set RSBusca = Nothing
    
End Function

Public Function CantUser() As Long
    Dim RsC As New ADODB.Recordset
            
    RsC.Open "SELECT * FROM Usuarios ", CN, adOpenStatic, adLockReadOnly
    
    CantUser = RsC.RecordCount
    
    RsC.Close
    Set RsC = Nothing
End Function

Public Function AgregarUsuario(USR As String, PSW As String) As Long
    'devuleve 0 en OK
    'devuelve 1 si ya existe el nombre de usuario
    'devuelve 2 si no cargo nada en alguno de los parametros
    
    If USR = "" Or PSW = "" Then
        AgregarUsuario = 2
        Exit Function
    End If
    
    If ExisteNCuenta("Usuario", "Usuarios", USR) > 0 Then
        AgregarUsuario = 1
        Exit Function
    End If
    
    CN.Execute "INSERT INTO Usuarios (Usuario,Contraseña) VALUES ('" + _
        USR + "','" + PSW + "')"
    
    'lo habilito ya para abrir y cerrar sesión (6)!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Dim tIdU As Long
    tIdU = GetID("Usuario", "Usuarios", USR)
    
    CN.Execute "INSERT INTO Habilitados (IdUsuario,IdEvento) VALUES (" + _
        CStr(tIdU) + ",6)"
    CN.Execute "INSERT INTO Habilitados (IdUsuario,IdEvento) VALUES (" + _
        CStr(tIdU) + ",5)"
    
    AgregarUsuario = 0
End Function

Public Function ModificarUsuario(exNombre As String, newNombre As String, _
    Optional exPSW As String, Optional newPSW As String) As Long
    'devuleve 0 en OK
    'devuelve 1 si ya existe el nombre de usuario
    'devuelve 2 si no cargo nada en alguno de los parametros
    'devuelve 3 por problemas con la clave
    
    'veo que haya cargado newnombre
    If newNombre = "" Then
        ModificarUsuario = 2
        Exit Function
    End If

    
    If ExisteNCuenta("Usuario", "Usuarios", newNombre) > 0 Then
        If exNombre = newNombre Then
            If exPSW = newPSW Then
                'nada escribio lo mismo, solo la descripcion
                Exit Function
            Else 'quiere cambiar la clave nomas
                If ValidarClave(exNombre, exPSW) = 0 Then
                    
                    'veo que haya cargado newPSW
                    If newPSW = "" Then
                        ModificarUsuario = 2
                        Exit Function
                    End If
                    
                    CN.Execute "UPDATE Usuarios SET Contraseña = '" + _
                        newPSW + "' WHERE Usuario = '" + exNombre + "'"
                    ModificarUsuario = 0
                Else
                    ModificarUsuario = 3
                End If
                
            End If
        Else
            ModificarUsuario = 1
        End If
        Exit Function
    End If
    
    Dim S As String
    S = "UPDATE Usuarios SET Usuario = '" + _
        newNombre + "' WHERE Usuario = '" + exNombre + "'"
    CN.Execute S
    
    ModificarUsuario = 0
End Function

    
Public Function ModificarEvento(exEvento As String, newEvento As String) As Long
    '0 OK
    '1 algo vacio
    '2 ya existe el nombre de esa cuenta
    
    '(1) que no sea vacio el nuevo nombre
    If newEvento = "" Or exEvento = "" Then
        ModificarEvento = 1
        Exit Function
    End If
    
    '(2) si no existe otro evento con ese nombre
    If ExisteNCuenta("Evento", "Eventos", newEvento) > 0 Then
        'veo que no sea el mismo
        If exEvento = newEvento Then
            'nada
        Else
            ModificarEvento = 2

        End If
        Exit Function
    End If
    
    'llego aca entonces ta joya
    CN.Execute "UPDATE Eventos SET Evento = '" + newEvento + "' WHERE " + _
        "Evento = '" + exEvento + "'"
    ModificarEvento = 0
End Function

Public Function ValidarClave(USR As String, PSW As String) As Long
    '0 OK,1 problema con usr,2 la clave esta mal
    Dim Rs As New ADODB.Recordset
    
    Rs.Open "SELECT * FROM Usuarios WHERE Usuario = '" + USR + "'", _
        CN, adOpenStatic, adLockReadOnly
    If Rs.RecordCount <> 1 Then
        ValidarClave = 1
    Else
        If PSW = Rs("Contraseña") Then
            ValidarClave = 0
        Else
            ValidarClave = 2
        End If
    End If
    
    Rs.Close
    Set Rs = Nothing
End Function

Public Function AgregarEvento(Evento As String) As Long
    '0 OK, 1 mal
    Dim UIDm1 As Long
    
    UIDm1 = UltIDEmasUno
    
    If ExisteNCuenta("Evento", "Eventos", Evento) > 0 Or Evento = "" Then
        AgregarEvento = 1
    Else
        CN.Execute "INSERT INTO Eventos (ID,Evento) VALUES (+" + _
            CStr(UIDm1) + ", '" + Evento + "')"
        AgregarEvento = 0
    End If
End Function

Private Function UltIDEmasUno() As Long
    Dim rsU As New ADODB.Recordset
    rsU.Open "SELECT TOP 1 ID FROM Eventos ORDER BY ID DESC", CN, adOpenStatic, adLockReadOnly
    
    UltIDEmasUno = rsU("ID") + 1
    
    rsU.Close
    Set rsU = Nothing
End Function

Public Sub EliminarEvento(Evento As String)
    CN.Execute "DELETE FROM Eventos WHERE Evento = '" + Evento + "'"
    'tambien se van a borrar movimientos y relaciones
End Sub

Public Sub EliminarUsuario(Usuario As String)
    CN.Execute "DELETE FROM Usuarios WHERE Usuario = '" + Usuario + "'"
    'tambien se van a borrar movimientos y relaciones
End Sub

Public Function RelacionarEvento(Evento As String, Usuario As String) As Long
    ',     UsuarioAdministrador As String, PswAdm As String) As Long
    '0 ok
    '1 mal la clave de administrador
    '2 usuario o evento no existen
    '3 no esta habilitado para el EVENTO asignar EVENTO
    '4 ya estaba la relacion
    
'    '(1) empiezo con clave de administrador
'    If ValidarClave(UsuarioAdministrador, PswAdm) > 0 Then
'        RelacionarEvento = 1
'        Exit Function
'    End If

    '(2) veo si existen usuarios y eventos
    If ExisteNCuenta("Usuario", "Usuarios", Usuario) = 1 And _
        ExisteNCuenta("Evento", "Eventos", Evento) = 1 Then
        'joiiiiia dejo que siga
    Else
        RelacionarEvento = 2
        Exit Function
    End If
    
'    '(3) ¿está habilitado para asignar eventos? 'id asignar eventos es si o si 1
'    Dim IDad As Long
'    IDad = GetID("Usuario", "Usuarios", UsuarioAdministrador)
'
'    If ExisteRelacion(IDad, 1) = 0 Then 'no existe la relacion con asig eventos
'        RelacionarEvento = 3
'        Exit Function
'    End If
    
    '(4) veo si no existìa antes la relación
    Dim IDU As Long, IDe As Long
    
    IDU = GetID("Usuario", "Usuarios", Usuario)
    IDe = GetID("evento", "eventos", Evento)
    
    If ExisteRelacion(IDU, IDe) > 1 Then
        If IDe = 1 Then
            RelacionarEvento = 4
            Exit Function
        Else
            RelacionarEvento = 0
        End If
    End If
    
    '(5) si llego aca listo HABILITO!!!!
    CN.Execute "INSERT INTO Habilitados (IdUsuario,IdEvento) VALUES (" + _
        CStr(IDU) + "," + CStr(IDe) + ")"
    
End Function

Public Sub EliminarRelacion(Evento As String, Usuario As String)
    'si no existe relacion no borra nada
    
    IDU = GetID("Usuario", "Usuarios", Usuario)
    IDe = GetID("evento", "eventos", Evento)
    
    CN.Execute "DELETE FROM Habilitados WHERE IdUsuario = " + CStr(IDU) + _
        " AND IdEvento = " + CStr(IDe)
End Sub

Public Function GetID(Campo As String, Tabla As String, Nombre As String) As Long
    '-1 si algo mal
    Dim rsB As New ADODB.Recordset
    
    rsB.Open "SELECT * FROM " + Tabla + " WHERE " + Campo + " = '" + _
        Nombre + "'", CN, adOpenStatic, adLockReadOnly
    
    If rsB.RecordCount <> 1 Then
        GetID = -1
    Else
        GetID = rsB("ID")
    End If
    
    rsB.Close
    Set rsB = Nothing
End Function

Public Function GetNombre(Campo As String, Tabla As String, ID As Long) As String
    Dim rsB As New ADODB.Recordset
    
    rsB.Open "SELECT * FROM " + Tabla + " WHERE ID = " + _
        CStr(ID), CN, adOpenStatic, adLockReadOnly
    
    If rsB.RecordCount <> 1 Then
        GetNombre = "Algo esta mal"
    Else
        GetNombre = rsB(Campo)
    End If
    
    rsB.Close
    Set rsB = Nothing
End Function

Public Function ExisteRelacion(IdUsuario As Long, IdEvento As Long) As Long
    'devuelve 0 si no existe
    'o devuelve el id de relacion
    
    'TOMO QUE SI HAY UN SOLO USUARIO ESTE HABILITADO PARA TODO!!!!!!!!!!!
    Dim RsUn As New ADODB.Recordset
    
    RsUn.Open "SELECT * FROM Usuarios", CN, adOpenStatic, adLockReadOnly
    
    If RsUn.RecordCount = 1 Then
        'Está habilitado para todo, mientras el evento exista
        If GetNombre("Evento", "Eventos", IdEvento) <> "Algo esta mal" Then
            ExisteRelacion = 1
        Else
            ExisteRelacion = 0
        End If
    Else
        Dim rsY As New ADODB.Recordset
        
        rsY.Open "SELECT * FROM Habilitados WHERE IdEvento = " + CStr(IdEvento), _
            CN, adOpenStatic, adLockReadOnly
        
        If rsY.RecordCount = 0 Then
            ExisteRelacion = 0
        Else
            rsY.MoveFirst
            Do While Not rsY.EOF
                'veo si de todos la relaciones este evento ya tiene este usuario
                If rsY("idusuario") = IdUsuario Then
                    ExisteRelacion = rsY("id")
                    Exit Do
                End If
                
                rsY.MoveNext
            Loop
        End If
        
        rsY.Close
        Set rsY = Nothing
    End If
    
    RsUn.Close
    Set RsUn = Nothing
End Function

Public Sub RegEvento(IdUsuario As Long, IdEvento As Long, _
    Optional Descripcion As String = "")
    
    CN.Execute "INSERT INTO MovUsuarios (Fecha,Hora,Minutos,IdUsuario,IdEvento," + _
        "Descripcion) VALUES (#" + stFechaSQL(Date) + "#," + _
        CStr(Hour(Now)) + "," + CStr(Minute(Now)) + "," + _
         CStr(IdUsuario) + "," + CStr(IdEvento) + ",'" + _
        Descripcion + "')"

End Sub

Public Function UltUsuarioIngresado(Optional IdEvento As Long = 0) As Long
    'da el ult usuario que entro a ese evento
    'si idEvento es 0 es que es el ultimo que entro para cualquier evento
    ' si da como resultado 0 es que es nadie
    Dim RSS As New ADODB.Recordset
    Dim S As String
    
    S = "SELECT TOP 1 IDUsuario FROM MovUsuarios"
    If IdEvento <> 0 Then S = S + " WHERE IdEvento = " + CStr(IdEvento)
    
    RSS.Open S + " ORDER BY ID desc", CN, adOpenStatic, adLockReadOnly
    
    If RSS.RecordCount = 0 Then
        UltUsuarioIngresado = 0
    Else
        UltUsuarioIngresado = RSS("idusuario")
    End If
    
    RSS.Close
    Set RSS = Nothing
End Function

Private Function stFechaSQL(Fecha As Date) As String
    stFechaSQL = CStr(Format(Fecha, "mm/dd/yyyy"))
End Function

Public Function Eventos(Optional IdUsuario As Long = 0) As String()
    'si idevento=0 muestra todos
    'si idusuario =-1 muestra mayores de 100
    
    Dim tmpEv() As String, i As Long, S As String
    Dim RsR As New ADODB.Recordset
    
    'pongo que tmpev(0)="" para asegurarme que algo le responde
    ReDim Preserve tmpEv(0)
    tmpEv(0) = ""
    
    S = "SELECT Evento FROM Eventos"
    If IdUsuario <> 0 Then
        If IdUsuario = -1 Then
            S = "SELECT Evento FROM Eventos WHERE ID >=100"
        Else
            S = "SELECT Eventos.Evento FROM Eventos INNER JOIN Habilitados " + _
                "ON Eventos.ID = Habilitados.IdEvento " + _
                "WHERE (((Habilitados.IdUsuario) = " + CStr(IdUsuario) + ")) " + _
                "GROUP BY Eventos.Evento"
        End If
    End If
    
    RsR.Open S, CN, adOpenStatic, adLockReadOnly
    
    i = 1
    
    If RsR.RecordCount > 0 Then
        RsR.MoveFirst
        Do While Not RsR.EOF
            ReDim Preserve tmpEv(i)
            tmpEv(i) = RsR("Evento")
            
            i = i + 1
            RsR.MoveNext
        Loop
    End If
    
    Eventos = tmpEv
End Function

Public Function Usuarios(Optional IdEvento As Long = 0) As String()
    'si idevento=0 muestra todos
    
    Dim tmpUSR() As String, i As Long, S As String
    Dim RsR As New ADODB.Recordset
    
    'pongo que tmpev(0)="" para asegurarme que algo le responde
    ReDim Preserve tmpUSR(0)
    tmpUSR(0) = ""
    
    S = "SELECT Usuario FROM Usuarios"
    
    If IdEvento <> 0 Then
        S = "SELECT Usuarios.Usuario FROM Usuarios INNER JOIN Habilitados " + _
            "ON Usuarios.ID = Habilitados.IdUsuario " + _
            "WHERE (((Habilitados.IdEvento) = " + CStr(IdEvento) + ")) " + _
            "GROUP BY Usuarios.Usuario"

    End If
    
    RsR.Open S, CN, adOpenStatic, adLockReadOnly
    
    i = 1
    
    If RsR.RecordCount > 0 Then
        RsR.MoveFirst
        Do While Not RsR.EOF
            ReDim Preserve tmpUSR(i)
            tmpUSR(i) = RsR("Usuario")
            
            i = i + 1
            RsR.MoveNext
        Loop
    End If
    
    Usuarios = tmpUSR
End Function

Public Sub DefinirPermisos()
    'frmAcomodar.AbrirDatos mDBFilename
    frmAcomodar.Show 1
End Sub

Public Function Ingresos(IdEvento As Long) As Long
    'Ingresos = frmINgreso.AbrirDatos(IdEvento, tmpDB)
    frmINgreso.Show 1
    '0 OK
    '1 Datos sin Cargar
    '2 Clave mal
    '3 Usuario no habilitado
    '4 no existe el usuario
    '5 cancelo el proceso
    
    'aca ya se cargo el usr y psw
    '(0) para empezar veo si cancelo
    If USR = "SOY DE LA T HASTA LA MUERTE" Then
        Ingresos = 5
        Exit Function
    End If
    
    '(1) Ver que haya cargado algo y validar
    If USR = "" Or PSW = "" Then
        Ingresos = 1
        Exit Function
    End If
    
    '(2) Veo si existe la cuenta
    If ExisteNCuenta("Usuario", "Usuarios", USR) = 0 Then
        Ingresos = 4
        Exit Function
    End If
    
    '(3) veo que contraseña ande
    If ValidarClave(USR, PSW) <> 0 Then
        Ingresos = 2
        Exit Function
    End If

    '(4) veo si usuario esta habilitado
    Dim IDU As Long
    IDU = GetID("Usuario", "Usuarios", USR)

    If ExisteRelacion(IDU, IdEvento) = 0 Then
        Ingresos = 3
        Exit Function
    End If

    'si llego aca esta OK, ya que esta registro el MovEvento
    Ingresos = 0
    RegEvento IDU, IdEvento
End Function

Public Function ValidarUsuario(IdEV As Long) As Long
    'hago que registre el usuario que entre
    '0 ok
    '-1 algo mal
    '-2 cancelo
    
    Dim i As Long, tmp As Long
   
    For i = 0 To 4
        tmp = Ingresos(IdEV)
            Select Case tmp
                Case 0
                    ValidarUsuario = 0
                    Exit For
                Case 1
                    MsgBox "Hay datos sin Cargar", vbExclamation, "Atención"
                Case 2
                    MsgBox "Contraseña Inválida", vbExclamation, "Atención"
                Case 3
                    MsgBox "Usuario no habilitado", vbExclamation, "Atención"
                Case 4
                    MsgBox "Usuario no existe", vbExclamation, "Atención"
                Case 5
                    ValidarUsuario = -2
                    Exit For
           End Select
    Next i
    
    If i >= 4 Then ValidarUsuario = -1
        
    
End Function

Public Function CambiarContrasena(IdUsuario As Long) As Long
    Dim Usuario As String
    Usuario = GetNombre("Usuario", "Usuarios", IdUsuario)
    
    'reinicio variables
    PSW = "": PSWm1 = "": PSWm2 = ""
    'abro el form
    frmPassword.Caption = frmPassword.Caption + " " + UCase(Usuario)
    frmPassword.Show 1
    
    'ya aca tengo los valores
    
    '0 OK
    '1 Datos sin Cargar
    '2 Clave mal
    '3 Mal confirmado
    
    '(1) veo si hay datos sin cargar
    If PSW = "" Or PSWm1 = "" Or PSWm2 = "" Then
        CambiarContrasena = 1
        Exit Function
    End If
    
    '(2) valido usuario
    If ValidarClave(Usuario, PSW) > 0 Then
        CambiarContrasena = 2
        Exit Function
    End If
    
    '(3) controlo que la confirmacion este bien
    If PSWm1 <> PSWm2 Then
        CambiarContrasena = 3
        Exit Function
    End If
    
    'supuestamente esta todo bien modifico clave nomas
    ModificarUsuario Usuario, Usuario, PSW, PSWm1
    CambiarContrasena = 0
        
End Function
    
Public Sub MostrarMovimientos()
    frmMovimientos.Show 1
End Sub

Public Sub LimpiarMov(DiasAtras As Long)
    CN.Execute "DELETE FROM MovUsuarios WHERE Fecha < #" + _
        stFechaSQL(Date - DiasAtras) + "#"
End Sub

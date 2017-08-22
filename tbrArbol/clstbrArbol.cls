VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clstbrArbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mArchivo As String
Private mConfVigentes() As String

'este archivo tiene renglones con el sig formato
'ID|IDANT|IDTXT|DESC|VALOR|Tipo
'45|3|ArchImgFondo2|Imagen de fondeo del formularios 2|3

'yipo vendria a ser el tipo de dato de valor
'si es archivo para poner un CMD que lo abra
'si es numero una caja de texto numerica

'TIPO ---------------------------------------------------------------------------------
'0 String
'1 Numero
'2 Fecha
'3 PathArchivo

'CONFIGURACIONES VIGENTES --(Stock n Cash) --------------------------------------------
'1|0|Titulo|0|tbrStock & Cash - Software Administración Pequeñas Empresas|0
'2|0|Dias Vencimiento|0|30|0
'3|0|IdSucursalPredeterminada|0|1|0
'4|0|Tipos IVA|0|0|0
'7|0|IVA Pred|0|21|0 (sin uso)
'8|0|LetraFacturaCompraPred|0|C|0 (sin uso)
'10|0|Calcular Vuelto||50|0
'11|0|Interes Mensual Pred||3|0
'12|0|Cantidad Maxima de Cuotas||36|0
'13|0|Prox Nro Factura||A-0002-303|0
'14|0|Vendedor Predeterminado|0|Larry O.|0
'17|0|Cantidad de Mov Productos||1|0
'20|0|Comisiones|0|0|0
'30|0|DatosTrabajoCliente|0|IDc*DetT*Ant*Fono*ingreso|0
'40|0|FormadePago|0|FIN|0
'40|0|MargenVenta|0|0|0

'ejemplo de hijos
'68|30|DTC 2||2*Pintor*1 año y 4 meses*no tiene*33,12|0
'69|30|DTC 3||3**20 años*903289243*10000|0
'70|40|FDP 3||CC*CUO*1200|0
'71|40|FDP -23||CC*CUO*30000|0
'72|40|FDP -20||CC*CUO*0|0
'65|20|Comision Jose||2|0
'66|20|Comision Larry O.||0,98|0
'73|30|DTC 1||1**23**333|0
'74|13|Factura X||0001-00000001|0
'67|13|Factura D||0001-60|0
'15|13|Factura A|0|0002-303|0
'44|50|MGV 2|0|33,33|0

Public Property Get Archivo() As String
    Archivo = mArchivo
End Property

Public Property Let Archivo(NewArch As String)
    mArchivo = NewArch
    OrdenarCFG
End Property

Public Property Get ConfVigentes(Index As Long) As String 'cada uno dice (ID*Titulo*Valor Predet)
    ConfVigentes = mConfVigentes(Index)
End Property

Public Property Let ConfVigentes(Index As Long, NewCV As String)
    ReDim Preserve mConfVigentes(Index)
    mConfVigentes(Index) = NewCV
End Property

Public Function SeguridadCopia()
    Dim FSO As New FileSystemObject
    Dim TmP As String
    
    TmP = FSO.GetSpecialFolder(TemporaryFolder)
    If Right(TmP, 1) <> "\" Then TmP = TmP + "\"

    'copia la configuracion ya que si se corta la luz se puede perder todo
    'ya paso, asi que es para hacer una copia cada vez que inicia y en el caso de que
    'haya un error evidente restauro
    
    FSO.CopyFile mArchivo, TmP + "config.abl"
    Set FSO = Nothing
End Function

Public Function SeguridadRestaura()
    Dim FSO As New FileSystemObject
    Dim TmP As String
    
    TmP = FSO.GetSpecialFolder(TemporaryFolder)
    If Right(TmP, 1) <> "\" Then TmP = TmP + "\"
    
    'RESTAURA la configuracion ya que si se corta la luz se puede perder todo
    'ya paso, asi que es para hacer una copia cada vez que inicia y en el caso de que
    'haya un error evidente restauro
    
    If FSO.FileExists(TmP + "config.abl") = True Then
        FSO.CopyFile TmP + "config.abl", mArchivo
    End If
    'si no esta no hace nada
    Set FSO = Nothing
End Function

Public Function AgregarNodo(IdAnterior As Long, Idtxt As String, _
    Descripcion As String, Valor As String, TipoDato As Long, _
    Optional IdNuevo As Long = -1) As Long
    'resultados posibles
    ' 0 Joya!
    ' -1 archivo = ""
    ' 1 ya existe una propiedad con ese nombre
    ' 2 hay datos sin completar
    
    If mArchivo = "" Then
        AgregarNodo = -1
        Exit Function
    End If
    
    If Idtxt = "" Then
        AgregarNodo = 2
        Exit Function
    End If
    
    If ExistePropiedad(Idtxt) <> 0 Then
        AgregarNodo = 1
        Exit Function
    End If
    
    'ubicar el ultimo numero disponible si no especifica el id
    Dim NewID As Long
    If IdNuevo = -1 Then
        NewID = GetLastID + 1
    Else
        NewID = IdNuevo
    End If
    
    Dim TE As TextStream
    Dim FSO As New Scripting.FileSystemObject
    
    If FSO.FileExists(mArchivo) = False Then FSO.CreateTextFile mArchivo, True
    
    Set TE = FSO.OpenTextFile(mArchivo, ForAppending, True)
    
    TE.WriteLine CStr(NewID) + "|" + CStr(IdAnterior) + "|" + Idtxt + "|" + _
        Descripcion + "|" + Valor + "|" + CStr(TipoDato)
        
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
    AgregarNodo = 0
        
End Function

Public Function ModificarNodo(IdBuscado As Long, Optional IdAnterior As Long = -1, _
    Optional Idtxt As String = "KKK", Optional Descripcion As String = "KKK", _
    Optional Valor As String = "KKK", Optional TipoDato As String = "KKK") As Long
    'pongo todo como opcional solo modifico la info elegida
    'empiezo armando un string que va a reemplazar el renglon modificado
    '---RESULTADOS POSIBLES----
    ' 0 - OK
    ' 1 - No se encontro propiedad -> que la agregue
    
    Dim Reemp As String, IDs() As String, TmpLinea() As String, C As Long
    Dim TE As TextStream, Encontrado As Boolean
    Dim FSO As New Scripting.FileSystemObject
    'veo como se va armando
    Dim IDa As Long, IDtx As String, Descr As String, Val As String, TDato As String
    
    'si no existe que lo agregue
    If ExistePropiedadID(IdBuscado) = "" Then
        AgregarNodo 0, CStr(Rnd * 1000), Descripcion, Valor, 0, IdBuscado
        'ahora si agrego las modificaciones
        ModificarNodo IdBuscado, IdAnterior, Idtxt, Descripcion, Valor, TipoDato
        Exit Function
    End If
    
    If IdAnterior = -1 Then
        IDa = GetInfo(IdBuscado, 1)
    Else
        IDa = IdAnterior
    End If
    
    If Idtxt = "KKK" Then
        IDtx = GetInfo(IdBuscado, 2)
    Else
        IDtx = Idtxt
    End If
    
    If Descripcion = "KKK" Then
        Descr = GetInfo(IdBuscado, 3)
    Else
        Descr = Descripcion
    End If
    
    If Valor = "KKK" Then
        Val = GetInfo(IdBuscado, 4)
    Else
        Val = Valor
    End If
    
    If TipoDato = "KKK" Then
        TDato = GetInfo(IdBuscado, 5)
    Else
        TDato = TipoDato
    End If
    'seria asi:....
    Reemp = CStr(IdBuscado) + "|" + CStr(IDa) + "|" + IDtx + "|" + Descr + "|" + _
        Val + "|" + CStr(TDato)
    
    'supongo que existe el archivo
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Encontrado = False
        
    'la mejor forma de modificar es
    '(1) copiar todo en un string(renglones), salvo el que se modifica
    '(2) al final de la matriz agregarle el nuevo
    '(3) meter todo en el archivo
    C = 0
    ReDim Preserve TmpLinea(0)
    TmpLinea(0) = ""
    
    Do While Not TE.AtEndOfStream
        C = C + 1
        ReDim Preserve TmpLinea(C)
        TmpLinea(C) = TE.ReadLine
        
        IDs = Split(TmpLinea(C), "|")
        If CLng(IDs(0)) = IdBuscado Then
            Encontrado = True
            C = C - 1 'para que borre el anterior
            ReDim Preserve TmpLinea(C)
        End If
    Loop
    ReDim Preserve TmpLinea(C + 1)
    TmpLinea(C + 1) = Reemp
    
    TE.Close
    
    Set TE = FSO.OpenTextFile(mArchivo, ForWriting, True)
    Dim OO As Long
    For OO = 1 To UBound(TmpLinea)
        TE.Close
        Set TE = FSO.OpenTextFile(mArchivo, ForAppending, True)

        TE.WriteLine TmpLinea(OO)
        'TE.SkipLine
    Next OO
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
    If Encontrado = False Then
        ModificarNodo = 1
    End If
       
    ModificarNodo = 0
End Function

Public Function EliminarNodo(IdBuscado As Long) As Long
    '---RESULTADOS POSIBLES----
    ' 0 - OK
    ' 1 - No se encontro propiedad
    
    Dim IDs() As String, TmpLinea() As String, C As Long
    Dim TE As TextStream, Encontrado As Boolean
    Dim FSO As New Scripting.FileSystemObject
    
    'supongo que existe el archivo
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Encontrado = False
        
    'la mejor forma de modificar es
    '(1) copiar todo en un string(renglones), salvo el que se modifica
    '(2) al final de la matriz agregarle el nuevo
    '(3) meter todo en el archivo
    
    C = 0
    ReDim Preserve TmpLinea(0)
    TmpLinea(0) = ""
    
    Do While Not TE.AtEndOfStream
        C = C + 1
        ReDim Preserve TmpLinea(C)
        TmpLinea(C) = TE.ReadLine
        
        IDs = Split(TmpLinea(C), "|")
        If CLng(IDs(0)) = IdBuscado Then
            Encontrado = True
            C = C - 1 'para que borre el anterior
            ReDim Preserve TmpLinea(C)
        End If
    Loop
    
    TE.Close
    
    Set TE = FSO.OpenTextFile(mArchivo, ForWriting, True)
    Dim OO As Long
    For OO = 1 To UBound(TmpLinea)
        TE.Close
        Set TE = FSO.OpenTextFile(mArchivo, ForAppending, True)

        TE.WriteLine TmpLinea(OO)
        'TE.SkipLine
    Next OO
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
    If Encontrado = False Then
        EliminarNodo = 1
    End If
   
    EliminarNodo = 0
End Function

Public Function ContarRenglones() As Long
    Dim TE As TextStream, FSO As New FileSystemObject, C As Long, F As String
    
    If FSO.FileExists(mArchivo) = False Then
        ContarRenglones = -1
        Set FSO = Nothing
        Exit Function
    End If
    
    'supongo que existe el archivo
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
        
    C = 0
    
    Do While Not TE.AtEndOfStream
        F = TE.ReadLine
        C = C + 1
    Loop
    
    ContarRenglones = C
    
    Set TE = Nothing
    Set FSO = Nothing
End Function

Public Function GetLastID() As Long
    'leer toooodo el archivo y buscar el ultimo
    If mArchivo = "" Then
        GetLastID = -1
        Exit Function
    End If
    
    Dim TE As TextStream
    Dim FSO As New Scripting.FileSystemObject
    
    If FSO.FileExists(mArchivo) = False Then
        FSO.CreateTextFile mArchivo, True
        GetLastID = 0
        Exit Function
    
    End If
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Dim TmpLinea As String, tmpID() As String, MaxId As Long
    Do While Not TE.AtEndOfStream
        TmpLinea = TE.ReadLine
        tmpID = Split(TmpLinea, "|")
        If CLng(tmpID(0)) > MaxId Then MaxId = CLng(tmpID(0))
    Loop
    
    GetLastID = MaxId
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
End Function

Public Function LeerRenglon(IDProp As Long) As String
    'leer toooodo el archivo y buscar el que tiene ese id
    If mArchivo = "" Then
        LeerRenglon = "No eligio archivo"
        Exit Function
    End If
    
    Dim TE As TextStream
    Dim FSO As New Scripting.FileSystemObject
    
    If FSO.FileExists(mArchivo) = False Then
        FSO.CreateTextFile mArchivo, True
        LeerRenglon = "No hay propiedad con ese ID"
        Exit Function
    
    End If
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Dim TmpLinea As String, tmpID() As String
    
    LeerRenglon = "No se encontro propiedad"
    
    Do While Not TE.AtEndOfStream
        TmpLinea = TE.ReadLine
        tmpID = Split(TmpLinea, "|")
        If CLng(tmpID(0)) = IDProp Then
            LeerRenglon = TmpLinea
            Exit Do
        End If
        
    Loop
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
End Function

Public Function ExistePropiedad(Nombre As String) As Long 'para idtxt
    'da 0 si esta OK (no existe)
    'si no da el nro del id que tiene ese idtxt
    
    'supongo que existe el archivo
    Dim TE As TextStream
    Dim FSO As New Scripting.FileSystemObject
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Dim TmpLinea As String, tmpIDtxT() As String
    
    ExistePropiedad = 0
    
    Do While Not TE.AtEndOfStream
        TmpLinea = TE.ReadLine
        tmpIDtxT = Split(TmpLinea, "|")
        If tmpIDtxT(2) = Nombre Then
            ExistePropiedad = CLng(tmpIDtxT(0))
            Exit Do
        End If
        
    Loop
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing

End Function

Public Function ExistePropiedadID(IdP As Long) As String
    'da "" si no existe
    'si no da el idtxt
    
    'supongo que existe el archivo
    Dim TE As TextStream
    Dim FSO As New Scripting.FileSystemObject
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Dim TmpLinea As String, tmpID() As String
    
    Do While Not TE.AtEndOfStream
        TmpLinea = TE.ReadLine
        tmpID = Split(TmpLinea, "|")
        If tmpID(0) = CStr(IdP) Then
            ExistePropiedadID = tmpID(2)
            Exit Do
        End If
        ExistePropiedadID = ""
    Loop
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing

End Function

Public Function GetHijos(IdBuscado As Long) As String()
    'devuelve una matriz con los valores DESDE 1 en adelante si tiene solo cero
    'es por que no hay nada
    
    Dim Ret() As String
    ReDim Ret(0) 'para que devuelva algo por mas que no haya registros
    
    'supongo que existe el archivo
    Dim TE As TextStream
    Dim FSO As New Scripting.FileSystemObject
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Dim TmpLinea As String, IdAnt() As String, C As Long
    
    C = 0
        
    Do While Not TE.AtEndOfStream
        TmpLinea = TE.ReadLine
        IdAnt = Split(TmpLinea, "|")
        If CLng(IdAnt(1)) = IdBuscado Then
            C = C + 1
            ReDim Preserve Ret(C)
            Ret(C) = IdAnt(0)
        End If
    Loop
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
    GetHijos = Ret
    
End Function

Public Function GetTodos() As String()
    'devuelve una matriz con los valores DESDE 1 en adelante si tiene solo cero
    'es por que no hay nada
    
    Dim Ret() As String
    ReDim Ret(0) 'para que devuelva algo por mas que no haya registros
    
    'supongo que existe el archivo
    Dim TE As TextStream
    Dim FSO As New Scripting.FileSystemObject
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Dim TmpLinea As String, IdAnt() As String, C As Long
    
    C = 0
        
    Do While Not TE.AtEndOfStream
        TmpLinea = TE.ReadLine
        IdAnt = Split(TmpLinea, "|")
            
        C = C + 1
        ReDim Preserve Ret(C)
        Ret(C) = IdAnt(0)
        
    Loop
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
    GetTodos = Ret
End Function

Public Function GetInfo(IDProp As Long, Index As Long) As String
    Dim Sp() As String
    
    Sp = Split(LeerRenglon(IDProp), "|")
    If UBound(Sp) = 0 Then
        GetInfo = "0"
    Else
        GetInfo = CStr(Sp(Index))
    End If

End Function

Public Function GetId(Idtxt As String) As Long
    '0 OK
    '-1 no esta el archivo
    '-2 no lo encontro
    
    Dim TE As TextStream
    Dim FSO As New Scripting.FileSystemObject
    
    If FSO.FileExists(mArchivo) = False Then
        FSO.CreateTextFile mArchivo, True
        GetId = -1
        Exit Function
    
    End If
    
    GetId = -2 'predeterminado
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Dim TmpLinea As String, tmpID() As String
    Do While Not TE.AtEndOfStream
        TmpLinea = TE.ReadLine
        tmpID = Split(TmpLinea, "|")
        If tmpID(2) = Idtxt Then
            GetId = CLng(tmpID(0))
        End If
    Loop
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
End Function

Public Sub TodosShow(Optional User As Boolean = True)
    frmPropiedades.AbrirDatos mArchivo, User
End Sub

Public Sub AgregarShow()
    frmAgregar.AbrirDatos mArchivo, -1, ""
End Sub

Private Sub OrdenarCFG()
    Dim TE As TextStream, Sp() As String, XX As Long
    Dim FSO As New Scripting.FileSystemObject
    On Local Error Resume Next
    
    If FSO.FileExists(mArchivo) = False Then
        'creo el archivo
        FSO.CreateTextFile mArchivo, True
        
        'le agrego las vigentes
        If UBound(mConfVigentes) > 0 Then
            For XX = 1 To UBound(mConfVigentes)
                Sp = Split(mConfVigentes(XX), "*")
                If ExistePropiedadID(CLng(Sp(0))) = "" Then AgregarNodo 0, _
                    Sp(1), "", Sp(2), 0, CLng(Sp(0))
            Next XX
        End If
        Exit Sub
    Else
        If ContarRenglones < 1 Then
            Set TE = FSO.OpenTextFile(mArchivo, ForWriting, True)
            TE.WriteLine "1|0|Nada|3|442|0"
            Exit Sub
        End If
    End If
    
'    '-- VEO LAS VIGENTES ?????? ------------------------------------------------------
'    If UBound(mConfVigentes) > 0 Then
'        For XX = 1 To UBound(mConfVigentes)
'            Sp = Split(mConfVigentes(XX), "*")
'            If ExistePropiedadID(CLng(Sp(0))) = "" Then AgregarNodo 0, _
'                Sp(1), "", Sp(2), 0, CLng(Sp(0))
'        Next XX
'    End If
    
    '-- LEO EL ARCHIVO ------------------------------------------------------
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
    Dim TmpLinea As String, Resp() As String, AA As Long
    AA = 0
    Do While Not TE.AtEndOfStream
        TmpLinea = TE.ReadLine
        'grabo en la matriz solo los que tengan al menos un | (supongo que lo demas
        'esta bien)
        If InStrRev(TmpLinea, "|") > 0 Then
            ReDim Preserve Resp(AA)
            Resp(AA) = TmpLinea
            AA = AA + 1
        End If
    Loop
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
    '-- PASO TODO POR ORDEN ------------------------------------------------------
    FSO.DeleteFile mArchivo, True
    FSO.CreateTextFile mArchivo
    Set TE = FSO.OpenTextFile(mArchivo, ForWriting, True)
    'TE.Write ""
    
    TE.Close
    Set TE = Nothing
    Set FSO = Nothing
    
    Dim BB As Long, C As Long, tmP2() As String
    
    AA = 0
    
    For BB = 1 To 1000
        For C = 0 To UBound(Resp)
            tmP2 = Split(Resp(C), "|")
            If BB = CLng(tmP2(0)) Then
                AgregarNodo CLng(tmP2(1)), tmP2(2), tmP2(3), tmP2(4), 0, CLng(tmP2(0))
                AA = AA + 1
            End If
        Next C
        If AA > UBound(Resp) Then Exit For
    Next BB
End Sub

Private Sub Class_Initialize()
    ReDim mConfVigentes(0)
    mConfVigentes(0) = "Nada"
End Sub

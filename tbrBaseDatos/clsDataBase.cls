VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mCN As New ADODB.Connection
Private mArchDB As String

Public Sub cn_CONECTAR_MDB(ArchivoMDB As String, Optional PSW As String = "zuliani")
    mCN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _
        ArchivoMDB + ";Jet OLEDB:Database Password= " + PSW
    
    mCN.Open
    
    mArchDB = ArchivoMDB
End Sub

Public Function CN_CLOSE() As Long
    If mCN.State = adStateOpen Then mCN.Close
    CN_CLOSE = mCN.State
End Function

Public Property Get CN() As ADODB.Connection
    Set CN = mCN
End Property

Public Function GetTop1Rs(Tabla As String, Campo As String, _
    Optional ORDERBY As String = "DESC", Optional WHERE As String = "NINGUNO", _
    Optional EsString As Boolean = False)
        
    'devuelve el top 1 de un campo determinado
    Dim RS As New ADODB.Recordset
    Dim tmpW As String
    
    If WHERE = "NINGUNO" Then
        tmpW = ""
    Else
        tmpW = " WHERE " + WHERE
    End If
    
    RS.CursorLocation = adUseClient
    RS.Open "Select TOP 1 " + Campo + _
        " FROM " + Tabla + tmpW + _
        " ORDER BY " + Campo + " " + ORDERBY, _
        mCN, adOpenStatic, adLockReadOnly
        
    If RS.RecordCount = 0 Then
        If EsString Then
            GetTop1Rs = ""
        Else
            GetTop1Rs = 0
        End If
        
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    If EsString Then
        GetTop1Rs = NoNuloS(RS.Fields(Campo))
    Else
        GetTop1Rs = NoNuloN(RS.Fields(Campo))
    End If
    
    RS.Close
    Set RS = Nothing
End Function

Public Sub ListarTabla(sSQL As String, Campo As String, LST_CMB As Object)
    Dim RS As New ADODB.Recordset
    RS.CursorLocation = adUseClient
    
    RS.Open sSQL, mCN, adOpenStatic, adLockReadOnly
    LST_CMB.Clear
    If RS.RecordCount = 0 Then Exit Sub
    RS.MoveFirst
    Do While Not RS.EOF
        LST_CMB.AddItem RS.Fields(Campo)
        RS.MoveNext
    Loop
    LST_CMB.ListIndex = 0
    RS.Close
    Set RS = Nothing
End Sub

Public Sub EXECUTE(sEXECUTE As String)
    mCN.EXECUTE sEXECUTE
End Sub

Public Function GetValInRS(Tabla As String, Campo As String, sWhere As String, _
    Optional EsString As Boolean = True)
    'devuelve el valor de un campo en una tabla en  un campo especificado dada
    'determinada condicion (supongo que es solo 1)
    
    Dim RS As New ADODB.Recordset
    RS.Open "Select " + Campo + " FROM " + Tabla + " WHERE " + sWhere, _
        mCN, adOpenStatic, adLockReadOnly
    If RS.RecordCount = 0 Then
        If EsString Then
            GetValInRS = ""
        Else
            GetValInRS = 0
        End If
    Else
        If EsString Then
            GetValInRS = NoNuloS(RS.Fields(Campo))
        Else
            GetValInRS = NoNuloN(RS.Fields(Campo))
        End If
    End If
    
    RS.Close
    Set RS = Nothing
End Function

Public Function ContarReg(sSQL As String) As Long
    'cuenta la cantidad de registros de un RS
    'sirve para duplicados por ejemplo
    Dim RS As New ADODB.Recordset
    RS.CursorLocation = adUseClient
    
    RS.Open sSQL, mCN, adOpenStatic, adLockReadOnly
    ContarReg = RS.RecordCount
    
    RS.Close
    Set RS = Nothing

End Function

Public Function SumarValInRS(Tabla As String, Campo As String, sWhere As String) As Single
    'devuelve la suma de un campo en una tabla en  un campo especificado dada
    'determinada condicion
    Dim Suma As Single
    Suma = 0
    
    Dim RS As New ADODB.Recordset
    RS.Open "Select " + Campo + " FROM " + Tabla + " WHERE " + sWhere, _
        mCN, adOpenStatic, adLockReadOnly
    If RS.RecordCount = 0 Then
        
    Else
        RS.MoveFirst
        Do While Not RS.EOF
            Suma = Suma + NoNuloN(RS.Fields(Campo))
            
            RS.MoveNext
        Loop
        
    End If
    
    SumarValInRS = Suma
    RS.Close
    Set RS = Nothing
End Function

Public Function SumarProducto(Tabla As String, Campo1 As String, Campo2 As String, _
    sWhere As String) As Single
    Dim Suma As Single
    Suma = 0
    
    Dim RS As New ADODB.Recordset
    RS.Open "SELECT " + Campo1 + ", " + Campo2 + " FROM " + Tabla + _
        " WHERE " + sWhere, mCN, adOpenStatic, adLockReadOnly
    If RS.RecordCount = 0 Then
        
    Else
        RS.MoveFirst
        Do While Not RS.EOF
            Suma = Suma + NoNuloN(RS(Campo1)) * NoNuloN(RS(Campo2))
            
            RS.MoveNext
        Loop
        
    End If
    
    SumarProducto = Suma
End Function

'Compact a JET database using ADO
'
'The following routines demonstrates how to compact a JET database using ADO:

'Purpose     :  Compact a JET (Access) database using ADO
'Inputs      :  sDatabasePath                   The path to the database path eg. C:\nwind.mdb
'               [bEncryptDatabase]              If True, encrypts the contents of the database
'Outputs     :  Returns zero if successful, else returns error code
'Notes       :  Requires "Microsoft Jet and Replication Objects X.X library",
'               where (X.X is greater than or equal to 2.1)
'               Compacts the database by creating a temporary database with the extension .tmp then,
'               if the compaction is successful, it overwrites the original database.
'               Will not work if anyone else is connected to the database.
'Revisions   :
'Assumptions :

Public Function CompactarBASE(sDatabasePath As String, PSW As String) As Long
    Dim oJRO As Object 'JRO.JetEngine
    Dim FSO As New Scripting.FileSystemObject

    On Error GoTo ErrFailed
    
    If Len(Dir$(sDatabasePath & ".tmp")) Then
        'Delete the existing temp database
        VBA.Kill sDatabasePath & ".tmp"
    End If
    
    Set oJRO = CreateObject("JRO.JetEngine")
    
    oJRO.CompactDatabase "Data Source=" & sDatabasePath + ";" + _
    "Jet OLEDB:Database Password=" + PSW, _
    "Data Source=" + sDatabasePath + ".tmp"
    
    'Delete the existing database
    VBA.Kill sDatabasePath
    'Rename the compacted database
    Name sDatabasePath & ".tmp" As sDatabasePath
    Set oJRO = Nothing
    
    CompactarBASE = 0
    Exit Function

ErrFailed:
    Debug.Print "Error al compactar base de datos: " & Err.Description
    CompactarBASE = Err.Number
    Set oJRO = Nothing
End Function

Public Function Backup(PathyArchivo As String, FldDestino As String) As Long
    '0 OK
    '1 No existe origen
    '2 No existe destino
    '3 Otro error
    
'    On Error GoTo OtroS
    Dim FSO As New Scripting.FileSystemObject
    
    If FSO.FileExists(PathyArchivo) = False Then
        Backup = 1
        Set FSO = Nothing
        Exit Function
    End If
    
    If FSO.FolderExists(FldDestino) = False Then
        Backup = 2
        Set FSO = Nothing
        Exit Function
    End If
    
    FSO.CopyFile PathyArchivo, FldDestino, True
    Backup = 0
    Set FSO = Nothing
    
    Exit Function
OtroS:
    Set FSO = Nothing
    Backup = 3
End Function

Public Function RestaurarBackup(PathyArchivo As String, FldDestino As String) As Long
    '0 Ok
    '1 No existe origen
    '2 No existe destino
    '3 cancelo
    '4 otro error
    
    On Error GoTo OtroS
    Dim FSO As New Scripting.FileSystemObject
    Dim nArch As String, DestArch As String
    
    nArch = Right(PathyArchivo, Len(PathyArchivo) - InStrRev(PathyArchivo, "\"))
    DestArch = FldDestino + nArch
    
    If FSO.FileExists(PathyArchivo) = False Then
        RestaurarBackup = 1
        Set FSO = Nothing
        Exit Function
    End If
    
    If FSO.FolderExists(FldDestino) = False Then
        RestaurarBackup = 2
        Set FSO = Nothing
        Exit Function
    End If
    
    If FSO.FileExists(DestArch) Then
        
        If MsgBox("¿Está seguro de reemplazar el archivo del " + _
            CStr(FileDateTime(DestArch)) + " de " + _
            CStr(FileLen(DestArch) / 1024) + " Kbytes " + vbCrLf + "por el Backup del " + _
            CStr(FileDateTime(PathyArchivo)) + " de " + _
            CStr(FileLen(PathyArchivo) / 1024) + " Kbytes?", _
            vbInformation + vbOKCancel) = vbCancel Then
            
            RestaurarBackup = 3
            Exit Function
        Else
            If CDate(FileDateTime(DestArch)) > CDate(FileDateTime(PathyArchivo)) Then
                If MsgBox("¿Está seguro de reemplazar a pesar de ser una base " + _
                    vbCrLf + "con fecha anterior?", vbInformation + vbOKCancel) = vbCancel Then
                    RestaurarBackup = 3
                    Exit Function
                End If
            
            End If
            
            'por las dudas que no restaure una base correcta
            FSO.CopyFile PathyArchivo, PathyArchivo + "Bup", True 'backup interno
            FSO.CopyFile PathyArchivo, FldDestino, True 'restaurado
            RestaurarBackup = 0
        End If
    Else
        RestaurarBackup = 1
        Set FSO = Nothing
        Exit Function
    End If
    
    Set FSO = Nothing
    Exit Function
OtroS:
    Set FSO = Nothing
    RestaurarBackup = 4

End Function


'generales-------------------------------------------------------------------------

Private Function NoNuloN(J) As Single
    If IsNumeric(J) Then
        NoNuloN = J
    Else
        NoNuloN = 0
    End If
End Function

Private Function NoNuloS(S) As String
    If IsNull(S) Then
        NoNuloS = ""
    Else
        NoNuloS = S
    End If
End Function

Public Function Tabla_Existe(La_Tabla As String) As Boolean
    Dim Resp As Boolean
    Resp = False
    On Local Error Resume Next
    
    Dim Datos() As String, I As Integer
    'Declara y abre el recordset
    Dim RS As Recordset
    
    Set RS = mCN.OpenSchema(adSchemaTables)
    
    Datos = Split(RS.GetString, vbTab)
    
    For I = 0 To UBound(Datos)
        'Compara el valor,( la tabla )
        If Trim(UCase(Datos(I))) = Trim(UCase(La_Tabla)) Then
           Resp = True ' La tabla existe
           Exit For
        End If
    Next
    
    Tabla_Existe = Resp
End Function



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl tbrBuscador 
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ScaleHeight     =   3600
   ScaleWidth      =   4500
   ToolboxBitmap   =   "tbrBuscador.ctx":0000
   Begin MSComctlLib.ListView LvBusca 
      Height          =   3015
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtBUSCA 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4350
   End
End
Attribute VB_Name = "tbrBuscador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mSqlSinLike As String
Private mCampoEnQueBuscar As String
Private mOrderBy As String
Private mSeparador As String
Private mCN As New ADODB.Connection
Private mArchivoMDB As String
Private mColumnas As String
Private mContrasena As String

Public Event Change()
Public Event Click()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Dim Demora As Single
Dim Demora2 As Single
Dim mDelay As Single

Public Sub SetDelayEventClick(newDelay As Single)
    mDelay = newDelay
End Sub

Private Sub lvBusca_Click()
    Demora = 0 'pone en cero el contador para lanar el evento
    RaiseEvent Click
    TryClick
End Sub

Private Sub TryClick()
    Demora2 = Timer
    DoEvents
    
    Do
        If Demora = 0 Then Exit Sub 'sale ya que hubo otro movimiento
        If Timer > Demora2 + mDelay Then Exit Do 'lanza elevento iaaa
    Loop
    
    RaiseEvent Click
End Sub

Public Sub Recargar()
    CN_Close
    ArchivoMDB = mArchivoMDB
    
    LvBusca.ColumnHeaders.Clear
    ColumnasTitulos
    
    txtBUSCA_Change
End Sub

Public Property Get ListCount()
    ListCount = LvBusca.ListItems.Count
End Property

Public Property Let SqlSinLike(NewSQL As String)
    mSqlSinLike = NewSQL
End Property

Public Property Get SqlSinLike() As String
    SqlSinLike = mSqlSinLike
End Property

Public Property Let ColumnasSepPorComasyParentesis(NewColumnas As String)
    'de esta forma fecha(ancho)/Detalle(ancho)/....
    mColumnas = NewColumnas
    ColumnasTitulos
End Property

Public Property Get ColumnasSepPorComasyParentesis() As String
    ColumnasSepPorComasyParentesis = mColumnas
End Property

Public Property Let Separador(NewSep As String)
    mSeparador = NewSep
End Property

Public Property Get Separador() As String
    Separador = mSeparador
End Property

Public Property Let Contrasena(NewPSW As String)
    mContrasena = NewPSW
End Property

Public Property Get Contrasena() As String
    Contrasena = mContrasena
End Property

Public Property Let CampoEnQueBuscar(NewCampo As String)
    mCampoEnQueBuscar = NewCampo
End Property

Public Property Get CampoEnQueBuscar() As String
    CampoEnQueBuscar = mCampoEnQueBuscar
End Property

 'va a hacer falta escribir "order by" igual anda de 10
Public Property Let OrderBy(NewOrder As String)
    mOrderBy = NewOrder
End Property

Public Property Get OrderBy() As String
    OrderBy = mOrderBy
End Property

Public Sub PonerFoco()
    txtBUSCA.SetFocus
End Sub

Private Function GetColumnas() As String()
    Dim SP() As String
    
    SP = Split(mColumnas, "/")
    
    GetColumnas = SP
End Function

Private Sub ColumnasTitulos()
    Dim I As Long, Col() As String, SP() As String
        
    LvBusca.ColumnHeaders.Clear
    
    Col = GetColumnas

    For I = 0 To UBound(Col)
        SP = Split(Col(I), "(")
        LvBusca.ColumnHeaders.Add , , SP(0), CLng(Left(SP(1), Len(SP(1)) - 1))
    Next I
End Sub

Private Sub LvBusca_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RaiseEvent Click
End Sub

Private Sub LvBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    'RaiseEvent Click
End Sub

Private Sub LvBusca_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtBUSCA_Change()
    Dim tmP As String
    Dim RSBUSCAR As New ADODB.Recordset
    Dim tmPN As Long, Buscar As String, Buscar2 As String
    Dim HH As Long, Campos() As String
        
    LvBusca.ListItems.Clear
    HH = 0: ReDim Campos(0): Buscar = "": Buscar2 = ""
       
    'pasa si el sql ya tiene un where
    If InStrRev(mSqlSinLike, "WHERE", , vbTextCompare) <> 0 Then
        tmP = " AND "
    Else
        tmP = " WHERE "
    End If
    
    If RSBUSCAR.State = adStateOpen Then RSBUSCAR.Close
    tmPN = 1
    
    Campos = Split(mCampoEnQueBuscar, ",")
     ' va a filtrar al que diga /b, si es el unico campo no hace falta DEBE SER STRING!!
    
    If UBound(Campos) = 0 Then
        Buscar = Campos(0)
    Else
        For HH = 0 To UBound(Campos)
            If Right(Campos(HH), 2) = "/b" Then
                Buscar = Left(Campos(HH), Len(Campos(HH)) - 2)
            End If
            'con v corta se puede poner una segunda condicion (por ejemplo:
            'si busca por pelicula original o nombre en español)
            If Right(Campos(HH), 2) = "/v" Then
                Buscar2 = Left(Campos(HH), Len(Campos(HH)) - 2)
                Exit For
            End If
        Next HH
    End If
    
    Dim Agregado As String
    
    If Buscar2 = "" Then
        Agregado = ""
    Else
        Agregado = "OR " + Buscar2 + " like '%" + txtBUSCA + "%' "
    End If
    
    RSBUSCAR.CursorLocation = adUseClient
    RSBUSCAR.Open mSqlSinLike + tmP + Buscar + " like '%" + txtBUSCA + _
        "%' " + Agregado + mOrderBy, mCN, adOpenStatic, adLockReadOnly
    
'            Ademas se le puede agregar al final
'        /n al final para indicar que es numero
'        /f para fechas
'        /$ para currency
'        predeterminado es string
    
    RSBUSCAR.Requery
        
    Dim AA As Long, SS As String
    If RSBUSCAR.RecordCount > 0 Then
        RSBUSCAR.MoveFirst
    
        Do While Not RSBUSCAR.EOF
            SS = ""
            LvBusca.ListItems.Add tmPN
            For AA = 0 To UBound(Campos)
                Dim Ult2 As String 'ultimos dos caracteres del campo
                Ult2 = Right(Campos(AA), 2)
                Dim NombreRealCampo As String
                Select Case Ult2
                    Case "/n"
                        NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                        SS = CStr(NoNuloN(RSBUSCAR(NombreRealCampo)))
                        
                        If AA = 0 Then
                            LvBusca.ListItems(tmPN).Text = SS
                        Else
                            LvBusca.ListItems(tmPN).SubItems(AA) = SS
                        End If
                        
                    Case "/f"
                        NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                        SS = CStr(RSBUSCAR(NombreRealCampo))
                    
                        If AA = 0 Then
                            LvBusca.ListItems(tmPN).Text = SS
                        Else
                            LvBusca.ListItems(tmPN).SubItems(AA) = SS
                        End If
                    
                    Case "/$"
                        NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                        SS = FormatCurrency(RSBUSCAR(NombreRealCampo), , , , vbFalse)
                        
                        If AA = 0 Then
                            LvBusca.ListItems(tmPN).Text = SS
                        Else
                            LvBusca.ListItems(tmPN).SubItems(AA) = SS
                        End If
                    Case "/b", "/v" 'es solo es que sirve de busqueda no hago nada raro
                        NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                        SS = CStr(RSBUSCAR(NombreRealCampo))
    
                        If AA = 0 Then
                            LvBusca.ListItems(tmPN).Text = SS
                        Else
                            LvBusca.ListItems(tmPN).SubItems(AA) = SS
                        End If
                        
                    Case Else
                        SS = NoNuloS(RSBUSCAR(Campos(AA)))
                        
                        If AA = 0 Then
                            LvBusca.ListItems(tmPN).Text = SS
                        Else
                            LvBusca.ListItems(tmPN).SubItems(AA) = SS
                        End If
                End Select
                'si no es el ultimo poner la barra separadora
            Next AA
            RSBUSCAR.MoveNext
            
            tmPN = tmPN + 1
        Loop
        
        If LvBusca.ListItems.Count <> 0 Then
            'LvBusca.SetFocus
            Set LvBusca.SelectedItem = LvBusca.ListItems(1)
        End If
            
        'If LvBusca.ListItems.Count <> 0 Then LvBusca.SelectedItem.Index = 0
    End If
    
    RSBUSCAR.Close
    Set RSBUSCAR = Nothing
    
    RaiseEvent Change
End Sub

Private Sub txtBUSCA_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode = vbKeyDown Then LvBusca.SetFocus
End Sub

 'para evitar la comilla simple
Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub UserControl_Initialize()
    mOrderBy = "" 'de entrada pongo que sea ""
    mSeparador = "/"
    txtBUSCA = ""
    mDelay = 0.5 'demora poredeterminada
    LvBusca.ListItems.Clear
End Sub

Public Property Let ArchivoMDB(NewMDB As String)
    mArchivoMDB = NewMDB
    
    If mCN.State = adStateClosed Then
        mCN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _
            mArchivoMDB + ";Jet OLEDB:Database Password= " + mContrasena
        mCN.Open
    End If
End Property

Public Sub CN_Close()
    If mCN.State = adStateOpen Then mCN.Close
End Sub

Private Sub UserControl_Resize()
    txtBUSCA.Top = 50
    txtBUSCA.Width = UserControl.Width - 100
    txtBUSCA.Left = 50
    LvBusca.Top = txtBUSCA.Top + txtBUSCA.Height
    LvBusca.Height = UserControl.Height - (LvBusca.Top)
    LvBusca.Width = UserControl.Width - 100
    LvBusca.Left = 50
End Sub

Public Function GetLstSel(Optional Indice As Long = 0, Optional Fila = 0) As String
    Dim mFila As Long
    
    If LvBusca.ListItems.Count = 0 Then GetLstSel = "": Exit Function
    
    mFila = LvBusca.SelectedItem.Index 'predeterminado
    
    If Fila > 0 And Fila <= LvBusca.ListItems.Count Then
        mFila = Fila
    End If
    
    If Indice = 0 Then
        GetLstSel = LvBusca.ListItems(mFila).Text
    Else
        GetLstSel = LvBusca.ListItems(mFila).SubItems(Indice)
    End If
    
End Function

Public Sub BorrarRenglon(Renglon As Long)
    If Renglon <= LvBusca.ListItems.Count Then
        LvBusca.ListItems.Remove (Renglon)
    End If
End Sub

Public Property Get Text() As String
    Text = txtBUSCA.Text
End Property

Public Property Let Text(NewText As String)
    txtBUSCA.Text = NewText
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H0&)
    txtBUSCA.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtBUSCA.SelLength = PropBag.ReadProperty("SelLength", 0)
    Set txtBUSCA.Font = PropBag.ReadProperty("FontT", Ambient.Font)
    Set LvBusca.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_Terminate()
    If mCN.State = adStateOpen Then mCN.Close
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H0&)
    Call PropBag.WriteProperty("SelStart", txtBUSCA.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtBUSCA.SelLength, 0)
    Call PropBag.WriteProperty("Fontt", txtBUSCA.Font, Ambient.Font)
    Call PropBag.WriteProperty("Font", LvBusca.Font, Ambient.Font)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtBUSCA,txtBUSCA,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = txtBUSCA.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtBUSCA.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtBUSCA,txtBUSCA,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = txtBUSCA.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtBUSCA.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lstBUSCA,lstBUSCA,-1,Font
Public Property Get Font() As Font
    Set Font = LvBusca.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set LvBusca.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get FontT() As Font
    Set FontT = txtBUSCA.Font
End Property

Public Property Set FontT(ByVal New_Font As Font)
    Set txtBUSCA.Font = New_Font
    PropertyChanged "FontT"
End Property





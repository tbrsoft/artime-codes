VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMovimientos 
   BackColor       =   &H007E7858&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos por Usuario"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbUsuarios 
      Height          =   360
      Left            =   4410
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1260
      Width           =   2205
   End
   Begin MSDataGridLib.DataGrid DGNoD 
      Height          =   4005
      Left            =   210
      TabIndex        =   2
      Top             =   1800
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   7064
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   465
      Left            =   10350
      TabIndex        =   1
      Top             =   6120
      Width           =   1005
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Usuario"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   1665
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMIENTOS POR USUARIO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   2790
      TabIndex        =   0
      Top             =   450
      Width           =   4095
   End
End
Attribute VB_Name = "frmMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsND As New ADODB.Recordset
Dim acC As New clsTbrAccesos

Private Sub cmbUsuarios_Click()
    Dim IDU As Long
    IDU = acC.GetID("Usuario", "Usuarios", cmbUsuarios)
    
    MostrarMov (IDU)
    AcomodarDG
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ListarUsuarios
End Sub

Private Sub MostrarMov(IdUsuario As Long)
    Dim Rs As String
    
    Rs = "SELECT MovUsuarios.ID, MovUsuarios.Fecha, MovUsuarios.Hora, MovUsuarios.Minutos, " + _
        "Eventos.Evento, MovUsuarios.Descripcion " + _
        "FROM Eventos INNER JOIN MovUsuarios ON Eventos.ID = MovUsuarios.IdEvento " + _
        "WHERE (((MovUsuarios.IdUsuario) = " + CStr(IdUsuario) + ")) " + _
        "GROUP BY MovUsuarios.ID, MovUsuarios.Fecha, MovUsuarios.Hora, MovUsuarios.Minutos, " + _
        "Eventos.Evento, MovUsuarios.Descripcion " + _
        "ORDER BY MovUsuarios.ID DESC, MovUsuarios.Fecha DESC , MovUsuarios.Hora DESC , " + _
        "MovUsuarios.Minutos DESC"
    
    If RsND.State = adStateOpen Then RsND.Close
    
    RsND.CursorLocation = adUseClient
    RsND.Open Rs, CN, adOpenStatic, adLockReadOnly
    
    Set DGNoD.DataSource = RsND
    
End Sub

Private Sub AcomodarDG()
    DGNoD.Columns("ID").Width = 0
    DGNoD.Columns("Fecha").Width = 1100
    DGNoD.Columns("Fecha").Alignment = dbgCenter
    DGNoD.Columns("Hora").Width = 700
    DGNoD.Columns("Hora").Alignment = dbgCenter
    DGNoD.Columns("Minutos").Width = 700
    DGNoD.Columns("Minutos").Alignment = dbgCenter
    DGNoD.Columns("Evento").Width = 2700
    DGNoD.Columns("Descripcion").Width = 5300
    
    DGNoD.RowHeight = 250
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set DGNoD.DataSource = Nothing
    
    RsND.Close
    Set RsND = Nothing
    Set acC = Nothing
End Sub

Private Sub ListarUsuarios()
    Dim tmpUSR() As String, i As Long

    cmbUsuarios.Clear

    tmpUSR = acC.Usuarios

    For i = 1 To UBound(tmpUSR)
        cmbUsuarios.AddItem tmpUSR(i)
    Next
    
    cmbUsuarios.ListIndex = 0
End Sub


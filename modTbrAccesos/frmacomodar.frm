VERSION 5.00
Begin VB.Form frmAcomodar 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accesos"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmacomodar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Agregar Habilitación"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3570
      TabIndex        =   12
      Top             =   5160
      Width           =   4905
      Begin VB.CommandButton cmdHabilitar 
         Caption         =   "Agregar Habilitacion"
         Height          =   555
         Left            =   780
         TabIndex        =   15
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Relacion a habilitarse (Seleccione Evento y Usuario)"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   300
         TabIndex        =   14
         Top             =   390
         Width           =   1845
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2130
         TabIndex        =   13
         Top             =   390
         Width           =   2265
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Eventos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6555
      Left            =   210
      TabIndex        =   8
      Top             =   570
      Width           =   3255
      Begin VB.CommandButton cmdEliminarE 
         Caption         =   "Eliminar Evento"
         Enabled         =   0   'False
         Height          =   380
         Left            =   930
         TabIndex        =   22
         Top             =   5880
         Width           =   1605
      End
      Begin VB.CommandButton cmdEventoM 
         Caption         =   "Modificar Evento"
         Enabled         =   0   'False
         Height          =   380
         Left            =   930
         TabIndex        =   19
         Top             =   4950
         Width           =   1605
      End
      Begin VB.CommandButton cmdEvento 
         Caption         =   "Agregar Evento"
         Enabled         =   0   'False
         Height          =   380
         Left            =   930
         TabIndex        =   10
         Top             =   5400
         Width           =   1605
      End
      Begin VB.ListBox lstEventos 
         Height          =   4140
         Left            =   150
         TabIndex        =   9
         Top             =   630
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Eventos"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   390
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4515
      Left            =   3540
      TabIndex        =   2
      Top             =   570
      Width           =   4935
      Begin VB.CommandButton cmdEliminarU 
         Caption         =   "Eliminar Usuario"
         Height          =   380
         Left            =   450
         TabIndex        =   21
         Top             =   3930
         Width           =   1545
      End
      Begin VB.CommandButton cmdKill 
         Caption         =   "Eliminar Habilitación"
         Height          =   380
         Left            =   2670
         TabIndex        =   20
         Top             =   2280
         Width           =   1755
      End
      Begin VB.CommandButton cmdNewNombre 
         Caption         =   "Modificar Nombre"
         Height          =   380
         Left            =   2640
         TabIndex        =   17
         Top             =   3780
         Width           =   1695
      End
      Begin VB.CommandButton cmdUsuario 
         Caption         =   "Agregar Usuario"
         Height          =   380
         Left            =   450
         TabIndex        =   5
         Top             =   3480
         Width           =   1545
      End
      Begin VB.ListBox lstUsuarios 
         Height          =   2700
         Left            =   180
         TabIndex        =   4
         Top             =   660
         Width           =   2175
      End
      Begin VB.ListBox lstUsuariosH 
         Height          =   1260
         Left            =   2490
         TabIndex        =   3
         Top             =   900
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Seleccionado"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2580
         TabIndex        =   18
         Top             =   2940
         Width           =   1665
      End
      Begin VB.Label lblUsuario 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2550
         TabIndex        =   16
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Todos"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   270
         TabIndex        =   7
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Habilitados para el evento seleccionado"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   2550
         TabIndex        =   6
         Top             =   360
         Width           =   1665
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   405
      Left            =   7410
      TabIndex        =   1
      Top             =   7380
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eventos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2460
      TabIndex        =   0
      Top             =   120
      Width           =   4965
   End
End
Attribute VB_Name = "frmAcomodar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim acC As New clsTbrAccesos

Private Sub cmdEliminarE_Click()
    If lstEventos.ListIndex = -1 Then Exit Sub
    
    If MsgBox("¿Está seguro de eliminar el evento " + UCase(lstEventos) + _
        "?", vbExclamation + vbYesNo, "Eliminar Evento") = vbNo Then Exit Sub
    
    acC.EliminarEvento lstEventos
    
    Actualizar
    
End Sub

Private Sub cmdEliminarU_Click()
    If lstUsuarios.ListIndex <= 0 Then Exit Sub
    
    If MsgBox("¿Está seguro de eliminar el usuario " + UCase(lstUsuarios) + _
        "?", vbExclamation + vbYesNo, "Eliminar Usuario") = vbNo Then Exit Sub
    
    acC.EliminarUsuario lstUsuarios
    
    Actualizar
End Sub

Private Sub cmdEvento_Click()
    Dim NwE As String, prU As Long
    
    NwE = InputBox("Ingrese nuevo evento", "Eventos")
    
    prU = acC.AgregarEvento(NwE)
    
    Select Case prU
        Case 0
            MsgBox "Se agrego " + UCase(NwE) + " como nuevo evento"
        Case 1
            MsgBox "Problemas", vbExclamation
    End Select
    
    Actualizar
End Sub

Private Sub cmdEventoM_Click()
    If lstEventos.ListIndex = -1 Then Exit Sub
    
    Dim NewEv As String
    NewEv = InputBox("Ingrese el Nuevo Nombre para este Evento", _
        "Nuevo Nombre Evento", lstEventos)
    
    Select Case acC.ModificarEvento(lstEventos, NewEv)
        Case 0
            Actualizar
        Case 1
            MsgBox "Nombre no válido", vbInformation, "Atención"
        Case 2
            MsgBox "Ya existe evento con ese nombre", vbInformation, "Atención"
    End Select
    
End Sub

Private Sub cmdHabilitar_Click()
    If lstEventos.ListIndex = -1 Or lstUsuarios.ListIndex = -1 Then Exit Sub
    
    Dim Hab As Long
    
    Hab = acC.RelacionarEvento(lstEventos, lstUsuarios)
        
    Select Case Hab
        Case 0
            MsgBox "Se agrego nueva habilitación" + vbCrLf + _
                Label3, vbInformation, "Registro Correcto"
        'Case 1
        '    MsgBox "Clave Administrador Incorrecta", vbExclamation, "Atención"
        Case 2
            MsgBox "Usuario o Evento No existen", vbExclamation, "Atención"
        'Case 3
         '   MsgBox "Administrador no habilitado", vbExclamation, "Atención"
        Case 4
            MsgBox "Habilitación ya existe", vbExclamation, "Atención"
        
    End Select
    
    Actualizar
End Sub

Private Sub cmdKill_Click()
    If lstUsuariosH.ListIndex = -1 Then
        MsgBox "No ha seleccionado ninguna habilitación." + vbCrLf + _
            "Seleccione una y luego presione nuevamente", vbInformation, "Atención"
        Exit Sub
    End If
    
    If lstEventos = "Asignar Eventos" And lstUsuariosH.ListCount = 1 Then
        MsgBox "Debe haber al menos un habilitado para este evento", vbInformation, "Atención"
        Exit Sub
    End If
    
    acC.EliminarRelacion lstEventos, lstUsuariosH
        
    Actualizar
End Sub

Private Sub cmdNewNombre_Click()
    If lstUsuarios.ListIndex = -1 Then Exit Sub
    'devuleve 0 en OK
    'devuelve 1 si ya existe el nombre de usuario
    'devuelve 2 si no cargo nada en alguno de los parametros
    'devuelve 3 por problemas con la clave
    
    Select Case acC.ModificarUsuario(lstUsuarios, _
        InputBox("Ingrese el nuevo Nombre", "Modificar Nombre", lstUsuarios))
        Case 0
            Actualizar
            'joya
        Case 1
            MsgBox "Ya existe un usuario con ese nombre", vbInformation, "Atención"
        Case 2
            MsgBox "Hay datos sin cargar", vbInformation, "Atención"
        Case 3
            MsgBox "Clave Incorrecta", vbInformation, "Atención"
        
    End Select
End Sub

Private Sub cmdUsuario_Click()
    Dim prU As Long, NwU As String
    
    NwU = InputBox("nuevo usuario")
    
    prU = acC.AgregarUsuario(NwU, InputBox("clave"))
    
    Select Case prU
        Case 0
            MsgBox "Se agrego " + UCase(NwU) + " como nuevo Usuario"
        Case 1
            MsgBox "Ya tiene usuario con ese nombre", vbExclamation
        Case 2
            MsgBox "Dejo en blanco uno de los campos", vbExclamation
    End Select
    
    Actualizar
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    Actualizar
End Sub

Private Sub ListarEventos()
    Dim tmpEv() As String, i As Long
    
    lstEventos.Clear
    
    tmpEv = acC.Eventos
    
    For i = 1 To UBound(tmpEv)
        lstEventos.AddItem tmpEv(i)
    Next
    
    lstEventos.ListIndex = 0
End Sub

Private Sub ListarUsuarios()
    Dim tmpUSR() As String, i As Long

    lstUsuarios.Clear

    tmpUSR = acC.Usuarios

    For i = 1 To UBound(tmpUSR)
        lstUsuarios.AddItem tmpUSR(i)
    Next
    
    lstUsuarios.ListIndex = 0
End Sub

Private Sub ListarUsuariosH(IDe As Long)
    Dim tmpUSR() As String, i As Long

    lstUsuariosH.Clear

    tmpUSR = acC.Usuarios(IDe)

    For i = 1 To UBound(tmpUSR)
        lstUsuariosH.AddItem tmpUSR(i)
    Next
    
    If UBound(tmpUSR) > 0 Then lstUsuariosH.ListIndex = 0
End Sub

Private Sub Actualizar()
    Dim i As Long
    Dim ixEv As Long, ixUs As Long
    
    ixEv = lstEventos.ListIndex
    ixUs = lstUsuarios.ListIndex
    If ixEv = -1 Then ixEv = 0
    If ixUs = -1 Then ixUs = 0
    
    ListarEventos
    ListarUsuarios

    'cargo administradores
    Dim tmpUSR() As String

    lstEventos_Click
    
    'veo usuario
    Dim ULog As String
    ULog = acC.GetNombre("Usuario", "Usuarios", acC.UltUsuarioIngresado)
    
    Label7 = "Usuario: " + ULog
    
    'hago que quede donde estaba el index
    If lstEventos.ListCount - 1 < ixEv Then
        If lstEventos.ListCount >= 0 Then lstEventos.ListIndex = ixEv - 1
    Else
        lstEventos.ListIndex = ixEv
    End If
    
    If lstUsuarios.ListCount - 1 < ixUs Then
        If lstUsuarios.ListCount >= 0 Then lstUsuarios.ListIndex = ixUs - 1
    Else
        lstUsuarios.ListIndex = ixUs
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 106 Then
        cmdEvento.Enabled = True
        cmdEventoM.Enabled = True
        cmdEliminarE.Enabled = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set acC = Nothing
End Sub

Private Sub lstEventos_Click()
    If lstEventos.ListIndex = -1 Or lstUsuarios.ListIndex = -1 Then Exit Sub
    
    Dim IdEV As Long
    IdEV = acC.GetID("Evento", "Eventos", lstEventos)
    
    ListarUsuariosH (IdEV)
    
    Label3 = UCase(lstEventos) + vbCrLf + "-" + vbCrLf + UCase(lstUsuarios)
End Sub

Private Sub lstUsuarios_Click()
    If lstEventos.ListIndex = -1 Or lstUsuarios.ListIndex = -1 Then Exit Sub
    
    Label3 = UCase(lstEventos) + vbCrLf + "-" + vbCrLf + UCase(lstUsuarios)
    lblUsuario = lstUsuarios
End Sub


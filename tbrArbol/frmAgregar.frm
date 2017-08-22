VERSION 5.00
Begin VB.Form frmAgregar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agregar Propiedad"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAgregar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbDependeDe 
      Height          =   360
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1740
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   435
      Left            =   2520
      TabIndex        =   6
      Top             =   3090
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   435
      Left            =   1110
      TabIndex        =   5
      Top             =   3090
      Width           =   1095
   End
   Begin VB.ComboBox cmbTipoDato 
      Height          =   360
      ItemData        =   "frmAgregar.frx":0442
      Left            =   1500
      List            =   "frmAgregar.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2310
      Width           =   2175
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   645
      Left            =   1500
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   930
      Width           =   2145
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   1500
      TabIndex        =   0
      Top             =   450
      Width           =   2145
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo de Dato"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   270
      TabIndex        =   9
      Top             =   2310
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Depende de"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   270
      TabIndex        =   8
      Top             =   1770
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   270
      TabIndex        =   7
      Top             =   1020
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   270
      TabIndex        =   2
      Top             =   480
      Width           =   1065
   End
End
Attribute VB_Name = "frmAgregar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AB As New clstbrArbol
Dim IdP As Long, IDa As Long, Val As String

Private Sub cmdOK_Click()
    Dim Res As Long
    
    If cmbDependeDe = "Nadie" Then
        IDa = 0
    Else
        IDa = AB.GetId(cmbDependeDe)
    End If
    
    If cmbTipoDato = "Fecha" And Not IsDate(Val) Then Val = "1-1-2000"
    
    If IdP = -1 Then 'AGREGAR!!
        If AB.ExistePropiedad(txtNombre) <> 0 Then
            MsgBox "Ya existe una propiedad con ese nombre"
            Exit Sub
        End If
    
        'agrego nomas
        Res = AB.AgregarNodo(IDa, txtNombre, txtDescripcion, Val, cmbTipoDato.ListIndex)
        If Res <> 0 Then MsgBox "Hubo errores de grabación -" + CStr(Res), _
            vbInformation, "Atención"
        
    Else 'MODIFICAR!!
        If AB.ExistePropiedad(txtNombre) = 0 Or AB.ExistePropiedad(txtNombre) = IdP Then
            'nada
        Else
            MsgBox "Ya existe una propiedad con ese nombre"
            Exit Sub
        End If
        
        Res = AB.ModificarNodo(IdP, IDa, txtNombre, txtDescripcion, Val, _
            cmbTipoDato.ListIndex)
        If Res <> 0 Then MsgBox "Hubo errores de grabación -" + CStr(Res), _
            vbInformation, "Atención"
    
    End If
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'cargo el combo DependeDe con todos las propiedades
    'le agrego una que diga "Nadie", si es para modificar tengo que borrarle la misma
    Dim Prop() As String
    
    Prop = AB.GetTodos
    cmbDependeDe.Clear
    
    For i = 1 To UBound(Prop)
        cmbDependeDe.AddItem AB.GetInfo(CLng(Prop(i)), 2)
    Next i
    
    cmbDependeDe.AddItem "Nadie"
    
    If IdP = -1 Then
        cmbTipoDato.ListIndex = 0
    Else
        Me.Caption = "Modificar Propiedad"
        'lleno la info
        txtNombre = AB.GetInfo(IdP, 2)
        txtDescripcion = AB.GetInfo(IdP, 3)
        
        cmbTipoDato.ListIndex = AB.GetInfo(IdP, 5)
        
        i = 0
        For i = 0 To cmbDependeDe.ListCount - 1
            If cmbDependeDe.List(i) = txtNombre Then cmbDependeDe.RemoveItem i
        Next i
        
        IDa = AB.GetInfo(IdP, 1)
        
    End If
    
    If IDa = 0 Then
        cmbDependeDe = "Nadie"
    Else
        cmbDependeDe = AB.GetInfo(IDa, 2)
    End If

End Sub

Public Sub AbrirDatos(Archivo As String, IDProp As Long, Valor As String, _
    Optional IdAnt As Long = 0)
    'si IDprop es -1 es nuevo
    IdP = IDProp
    IDa = IdAnt
    Val = Valor
    
    AB.Archivo = Archivo
    
    Me.Show 1
End Sub

Private Sub txtDescripcion_GotFocus()
    txtDescripcion.SelStart = 0
    txtDescripcion.SelLength = Len(txtDescripcion)
End Sub

Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(txtNombre)
End Sub

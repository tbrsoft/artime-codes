VERSION 5.00
Begin VB.Form frmINgreso 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clave de Acceso"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmINgreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   1770
      TabIndex        =   5
      Top             =   1530
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   300
      TabIndex        =   4
      Top             =   1530
      Width           =   1185
   End
   Begin VB.TextBox txtPSW 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   210
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtUSR 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   900
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de usuario"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   750
      TabIndex        =   1
      Top             =   30
      Width           =   1785
   End
End
Attribute VB_Name = "frmINgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsACC As New clsTbrAccesos

Private Sub cmdCancelar_Click()
    USR = "SOY DE LA T HASTA LA MUERTE"
    Unload Me
End Sub

Private Sub Command1_Click()
    USR = txtUSR
    PSW = txtPSW

    Unload Me
End Sub

Private Sub Form_Load()
    If clsACC.GetNombre("Usuario", "Usuarios", 1) = "Administrador" And _
        clsACC.ValidarClave("Administrador", "1") = 0 And _
        clsACC.CantUser = 1 Then
            'nunca cambio nada desde que lo instalo
        
        txtUSR = "Administrador"
        txtPSW = "1"
    Else
        txtUSR = ""
        txtPSW = ""

    End If
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'clsACC.Desconectar 'no porque la dll sigue funcionando
    Set clsACC = Nothing
End Sub

Private Sub txtUSR_GotFocus()
    txtUSR.SelStart = 0
    txtUSR.SelLength = Len(txtUSR)
End Sub

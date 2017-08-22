VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   2640
      TabIndex        =   4
      Top             =   1740
      Width           =   1215
   End
   Begin VB.TextBox txtPSWm 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2100
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   180
      Width           =   1515
   End
   Begin VB.TextBox txtNewPSW1 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2100
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   1515
   End
   Begin VB.TextBox txtNewPSW2 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2100
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton cmdNewPSW 
      Caption         =   "Modificar Contraseña"
      Default         =   -1  'True
      Height          =   405
      Left            =   540
      TabIndex        =   3
      Top             =   1740
      Width           =   1905
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   810
      TabIndex        =   7
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Contraseña"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   540
      TabIndex        =   6
      Top             =   690
      Width           =   1515
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirme Contraseña"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   1935
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNewPSW_Click()
    PSW = txtPSWm
    PSWm1 = txtNewPSW1
    PSWm2 = txtNewPSW2

    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


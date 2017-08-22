VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCierresViejos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro Cierres"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCierresViejos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstResumen 
      Height          =   3180
      Left            =   390
      TabIndex        =   18
      Top             =   2760
      Width           =   4605
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Últimos Cierres"
      Height          =   435
      Left            =   5910
      TabIndex        =   17
      Top             =   5520
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   435
      Left            =   7830
      TabIndex        =   2
      Top             =   5500
      Width           =   1035
   End
   Begin VB.ComboBox cmbCierres 
      Height          =   360
      Left            =   2850
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   1425
   End
   Begin MSComCtl2.DTPicker txtFecha 
      Height          =   375
      Left            =   2850
      TabIndex        =   0
      Top             =   870
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   38854
   End
   Begin VB.Line Line1 
      X1              =   7620
      X2              =   8940
      Y1              =   4530
      Y2              =   4530
   End
   Begin VB.Label lblDif 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7700
      TabIndex        =   16
      Top             =   3930
      Width           =   1155
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sobrante de Caja"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5400
      TabIndex        =   15
      Top             =   3990
      Width           =   2145
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos de Caja"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1110
      TabIndex        =   14
      Top             =   2400
      Width           =   2145
   End
   Begin VB.Label lblHora 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "11:00 am"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2850
      TabIndex        =   13
      Top             =   1800
      Width           =   1365
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hora de Cierre"
      Height          =   345
      Left            =   810
      TabIndex        =   12
      Top             =   1800
      Width           =   1965
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Variación de Caja"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5400
      TabIndex        =   11
      Top             =   3390
      Width           =   2145
   End
   Begin VB.Label lblMov 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7695
      TabIndex        =   10
      Top             =   3360
      Width           =   1155
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo Al Cierre"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5400
      TabIndex        =   9
      Top             =   4710
      Width           =   2145
   End
   Begin VB.Label lblEF 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7695
      TabIndex        =   8
      Top             =   4650
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo Anterior"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5400
      TabIndex        =   7
      Top             =   2850
      Width           =   2145
   End
   Begin VB.Label lblEI 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7695
      TabIndex        =   6
      Top             =   2820
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID del Cierre"
      Height          =   345
      Left            =   750
      TabIndex        =   5
      Top             =   1350
      Width           =   1965
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Fecha"
      Height          =   345
      Left            =   1230
      TabIndex        =   4
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cierres Registrados"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1170
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmCierresViejos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnB As New clsPruebaContab
Dim mMdb As String, mPsW As String ' para abrir ultimos

Public Function AbrirDatos(MDB As String, PSW As String)
    cnB.ArchMDB = MDB
    cnB.PSW = PSW
    cnB.Conectar
    
    mMdb = MDB
    mPsW = PSW
    
    Me.Show 1
End Function

Private Sub cmbCierres_Click()
    MostrarResumen
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    frmUltimos.AbrirDatos mMdb, mPsW
End Sub

Private Sub Form_Load()
    txtFecha = Date
    txtFecha_Change
End Sub

Private Sub MostrarResumen()
    If cmbCierres.ListIndex = -1 Then Exit Sub
    
    lstResumen.Clear
    lblHora = ""
    
    Dim stResumen As String, Resumen() As String, I As Long, EF As Single
    Dim Dif As Single, Var As Single

    lblHora = cnB.NoNuloS((cnB.GetValInRS("CierresViejos", "Hora", "IdCierre = " + cmbCierres)))
    stResumen = cnB.GetCierre(CLng(cmbCierres.Text))
    Resumen = Split(stResumen, "\\")

    For I = 0 To UBound(Resumen)
        lstResumen.AddItem CStr(Resumen(I))
    Next I
    
    'tmp = Split(CNB.GetCierre2(CLng(cmbCierres.Text)), "/")
    EF = cnB.NoNuloN(cnB.GetValInRS("CierresViejos", "Efvo", _
        "IdCierre = " + cmbCierres))
    Var = cnB.NoNuloN(cnB.GetValInRS("CierresViejos", "Variacion", _
        "IdCierre = " + cmbCierres))
    
    lblEF = FormatCurrency(EF)
    lblMov = FormatCurrency(Var)
    
    'diferencia de caja
    Dif = cnB.NoNuloN(cnB.GetValInRS("CierresViejos", "Diferencia", _
        "IdCierre = " + cmbCierres))
    
    If Dif < 0 Then
        lblDif.BackColor = &H8080FF
        Label9 = "Faltante de Caja"
    Else
        lblDif.BackColor = &HC0FFC0
        Label9 = "Sobrante de Caja"
    End If
    
    lblDif = FormatCurrency(Dif)
        
    'este lo saco por diferencia para no complicarlo
    lblEI = FormatCurrency(EF - Var - Dif)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cnB.CN_CLOSE
    Set cnB = Nothing
End Sub

Private Sub txtFecha_Change()
    cnB.CargarCombo cmbCierres, "SELECT IdCierre FROM CierresViejos WHERE " + _
        "Fecha = #" + cnB.stFechaSQL(txtFecha) + "# ORDER BY IDCierre DESC", _
        "IdCierre/n"
    
    If cmbCierres = "" Then
        lstResumen.Clear
        lblEI = FormatCurrency(0)
        lblDif = lblEI
        lblDif.BackColor = &HC0FFC0
        lblMov = lblEI
        lblEF = lblEI
    End If
    
End Sub

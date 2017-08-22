VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPropiedades 
   BackColor       =   &H00808080&
   Caption         =   "Configuraciones"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPropiedades.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker txtValorF 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   6420
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20905985
      CurrentDate     =   38916
   End
   Begin VB.TextBox txtValor 
      Height          =   345
      Left            =   930
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   6420
      Width           =   6120
   End
   Begin VB.ListBox lstCuentas 
      Height          =   5100
      Left            =   90
      TabIndex        =   5
      Top             =   480
      Width           =   8085
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   405
      Left            =   8370
      TabIndex        =   4
      Top             =   2130
      Width           =   975
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   405
      Left            =   8370
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   405
      Left            =   8490
      TabIndex        =   2
      Top             =   6390
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   345
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   6000
      Width           =   7215
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "Asignar"
      Height          =   375
      Left            =   7230
      TabIndex        =   0
      Top             =   6420
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Propiedades"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   180
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6450
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   6030
      Width           =   645
   End
End
Attribute VB_Name = "frmPropiedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AB As New clstbrArbol
Dim mArch As String
Dim IdPropSel As Long

Private Sub cmdAgregar_Click()
    If lstCuentas.ListIndex = -1 Then
        frmAgregar.AbrirDatos mArch, -1, ""
    Else
        frmAgregar.AbrirDatos mArch, -1, "", IdPropSel
    End If
    
    Reiniciar
End Sub

Private Sub cmdAsignar_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim TipoD As Long, ValorNuevo As String, F As String
    
    TipoD = AB.GetInfo(IdPropSel, 5)
    F = ""
    
    Select Case TipoD
        Case 0, 1 'numero - string
            ValorNuevo = txtValor
        
        Case 2 'fecha
            ValorNuevo = CStr(txtValorF)
            
        Case 3 ' Path Archivo
            Dim CD As New CommonDialog
            CD.FileName = txtValor
            CD.ShowOpen
        
            F = CD.FileName
            
            If F = "" Then Exit Sub
            ValorNuevo = F
        
    End Select
    
    AB.ModificarNodo IdPropSel, , , , ValorNuevo
    
    MsgBox "Se asignó -" + ValorNuevo + "- a " + AB.GetInfo(IdPropSel, 2), _
        vbInformation, "Modificación exitosa"
    
    Reiniciar
    
End Sub

Private Sub cmdModificar_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    If AB.GetInfo(IdPropSel, 5) = 3 Then 'fecha
        frmAgregar.AbrirDatos mArch, IdPropSel, CStr(txtValorF)
    Else
        frmAgregar.AbrirDatos mArch, IdPropSel, txtValor
    End If
End Sub


Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Reiniciar
End Sub

Private Sub Form_Load()
    'saco las propiedades principales
    Reiniciar
    
End Sub

Private Sub lstCuentas_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim Sp() As String, TipoD As Long
    
    Sp = Split(lstCuentas, ".")
    IdPropSel = CLng(Sp(1))
        
    'veo el tipo de dato para esconder el que no corresponda
    TipoD = CLng(AB.GetInfo(IdPropSel, 5))
    If TipoD = 2 Then 'fecha
        cmdAsignar.Left = 2550
        txtValor.Visible = False
        txtValorF.Visible = True
        txtValorF = CDate(AB.GetInfo(IdPropSel, 4))
    Else
        cmdAsignar.Left = 7230
        txtValor.Visible = True
        txtValorF.Visible = False
        txtValor = AB.GetInfo(IdPropSel, 4)
    End If
        
    txtDescripcion = AB.GetInfo(IdPropSel, 3)
    
End Sub

Private Sub lstCuentas_DblClick()
    'mostrar los subniveles del elegido
    Dim tmpIx As Long
    tmpIx = lstCuentas.ListIndex
    
    Dim CTA As Long, CTAS() As String, Sp() As String
    
    Sp = Split(lstCuentas, ".")
    IdPropSel = CLng(Sp(1))
    CTA = CLng(Sp(1))
    CTAS = AB.GetHijos(CTA)
    
    If UBound(CTAS) = 0 Then
        MsgBox "No tiene subcuentas!"
        Exit Sub
    End If
    
    'veo si ya mostro las subcuentas, si es asi las escondo
    If Right(lstCuentas, 1) = "*" Then
        'primero le borro el asterisco
        lstCuentas.List(tmpIx) = Left(lstCuentas.List(tmpIx), _
            Len(lstCuentas.List(tmpIx)) - 1)
        
        'escondo las subcuentas
        Dim Niv As Long, i As Long
        Niv = nNivel(tmpIx)
        
        i = tmpIx + 1
        
        Do While Not Niv >= nNivel(i)
            lstCuentas.RemoveItem i
            
            If i > lstCuentas.ListCount - 1 Then Exit Do
            'se baja por la eliminacion
            'I = I + 1
        Loop
        
        Exit Sub
    End If
    
    Dim A As Long
    For A = 1 To UBound(CTAS)
              
        'ponerle los mismos espacios que tenia mas 3
        lstCuentas.AddItem "   " + Sp(0) + "." + CTAS(A) + "." + _
            AB.GetInfo(CLng(CTAS(A)), 2), lstCuentas.ListIndex + 1
        
    Next A
    
    'marco que ya lo abrio
    lstCuentas.List(tmpIx) = lstCuentas.List(tmpIx) + "*"
End Sub

Private Function nNivel(IndiceLista As Long) As Long
    'veo el nivel de la cuenta seleccionada del listbox

    If IndiceLista = -1 Then
        nNivel = -1
        Exit Function
    End If
    
    Dim Spp() As String
    Spp = Split(lstCuentas.List(IndiceLista), ".")
    'spp(0) tiene los espacios que me van a decir en que nivel esta
    If Len(Spp(0)) = 0 Then  'es el nivel1
        nNivel = 1
    Else 'hago la formula, tiene que dar un numero redondo
        nNivel = Round(Len(Spp(0)) / 3 + 1, 0)
    End If
    
End Function

Private Sub Reiniciar()
    'saco las propiedades principales
    Dim Princ() As String
    
    Princ = AB.GetHijos(0)
    lstCuentas.Clear
    
    For i = 1 To UBound(Princ)
        lstCuentas.AddItem "." + CStr(AB.GetInfo(CLng(Princ(i)), 0)) + "." + _
            AB.GetInfo(CLng(Princ(i)), 2)
    Next i
    
    If lstCuentas.ListCount > 0 Then lstCuentas.ListIndex = 0

End Sub

Public Sub AbrirDatos(Archivo As String, Optional Usuario As Boolean = True)
    mArch = Archivo
    AB.Archivo = mArch
    
    If Usuario = True Then
        cmdAgregar.Visible = False
        cmdModificar.Visible = False
    End If
    Me.Show 1
End Sub


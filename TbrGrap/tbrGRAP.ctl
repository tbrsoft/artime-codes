VERSION 5.00
Begin VB.UserControl tbrGRAP 
   BackColor       =   &H00000000&
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ScaleHeight     =   3810
   ScaleWidth      =   4950
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   90
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lTIT 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   150
      TabIndex        =   3
      Top             =   30
      Width           =   4695
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L2"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   150
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L1"
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   3180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line EJEY 
      BorderColor     =   &H00FFFFFF&
      X1              =   750
      X2              =   750
      Y1              =   3450
      Y2              =   150
   End
   Begin VB.Line EJEX 
      BorderColor     =   &H00FFFFFF&
      X1              =   270
      X2              =   4560
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Shape SH 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      Height          =   1695
      Index           =   0
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "tbrGRAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private CB As Long 'cantidad de barras activas
Private mCampoString() As String
Private mCampoNum() As Single

'para graficar bien
Dim sMax As Single
Dim sMin As Single

Private mTitulo As String

Private Sub UserControl_Initialize()
    CB = 0
    ReDim mCampoString(0)
    ReDim mCampoNum(0)
End Sub

Public Property Get Titulo() As String
    Titulo = mTitulo
End Property

Public Property Let Titulo(newTitulo As String)
    mTitulo = newTitulo
    lTIT.Caption = mTitulo
End Property

Private Sub UserControl_Resize()
    RSZ
End Sub

Private Sub RSZ()
    L(0).Top = UserControl.Height - L(0).Height
    EJEX.Y1 = L(0).Top
    EJEX.Y2 = EJEX.Y1
    'estirando ejes
    EJEX.X2 = UserControl.Width
    EJEY.Y2 = 150
    EJEY.Y1 = UserControl.Height
    'anchos proporcionales
    L(0).Width = (UserControl.Width - EJEY.X1) / (CB + 1)
    List1.AddItem CStr(L(0).Width)
    SH(0).Width = (L(0).Width - 30)
    'reacomodar los derivados
    Dim B As Long
    For B = 1 To CB
        L(B).Top = L(0).Top
        L(B).Left = EJEY.X1 + (L(0).Width * (B - 1))
        L(B).Width = L(0).Width
        SH(B).Left = 30 + L(B).Left
        SH(B).Width = SH(0).Width
    Next B
    lTIT.Left = 0
    lTIT.Width = UserControl.Width
    lTIT.Alignment = 2
    
    
End Sub

Public Sub Descargar()
    On Local Error Resume Next
    Dim B As Long
    For B = 1 To CB
        Unload L(B)
        Unload SH(B)
    Next B
    
    ReDim Preserve mCampoString(0)
    ReDim Preserve mCampoNum(0)
End Sub

Public Sub ConectarRS(R As ADODB.Recordset, campoSTR As String, CampoNum As String, _
    ByRef MinValueFind As Single, MaxValueFind As Single)
    'tomo el recordset, lo transformo a mi gusto y despues se puede cerra del _
        otro lado, yano lo necesito
    
    'campoString es el de las etiqueta y el numerico de los valores
    
    R.MoveFirst: CB = 1
    Do While Not R.EOF
        If CB = 1 Then
            sMax = CSng(R.Fields(CampoNum))
            sMin = CSng(R.Fields(CampoNum))
        End If
        ReDim Preserve mCampoString(CB): mCampoString(CB) = R.Fields(campoSTR)
        ReDim Preserve mCampoNum(CB): mCampoNum(CB) = CSng(R.Fields(CampoNum))
        'ver max min
        If CSng(R.Fields(CampoNum)) > sMax Then sMax = CSng(R.Fields(CampoNum))
        If CSng(R.Fields(CampoNum)) < sMin Then sMin = CSng(R.Fields(CampoNum))
        
        Load L(CB): Load SH(CB)
        CB = CB + 1
        R.MoveNext
    Loop
    CB = CB - 1
    
    MinValueFind = sMin
    MaxValueFind = sMax
    
    'pueden cerralo si quieren al rs
    RSZ
    
End Sub
    
Public Sub Mostrar()
    RSZ 'todas las dimensiones horizaontales de ancho de cada barra
    
    'ver el alto maximo de cada barra
    Dim topBarra As Single
    topBarra = EJEX.Y1 - 60
    
    Dim B As Long
    For B = 1 To CB
        L(B).Caption = mCampoString(B) + " = " + CStr(mCampoNum(B))
        'posicionar y dimensionar
        SH(B).Height = (mCampoNum(B) * topBarra) / sMax
        
        SH(B).Top = EJEX.Y1 - SH(B).Height
        Randomize
        SH(B).BackColor = Int(Rnd * 60000)
        'mostrar
        L(B).Visible = True
        SH(B).Visible = True
    Next B
End Sub

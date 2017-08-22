VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUltimos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Últimos Cierres"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUltimos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   435
      Left            =   4530
      TabIndex        =   5
      Top             =   6000
      Width           =   1125
   End
   Begin VB.ComboBox cmbNCierres 
      Height          =   360
      ItemData        =   "frmUltimos.frx":0442
      Left            =   3750
      List            =   "frmUltimos.frx":0458
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   900
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   405
      Left            =   9000
      TabIndex        =   1
      Top             =   6000
      Width           =   1005
   End
   Begin MSDataGridLib.DataGrid DGC 
      Height          =   4245
      Left            =   90
      TabIndex        =   0
      Top             =   1350
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   7488
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione cantidad de Cierres"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   900
      TabIndex        =   4
      Top             =   960
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Últimos 10 Cierres"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2220
      TabIndex        =   3
      Top             =   270
      Width           =   4455
   End
End
Attribute VB_Name = "frmUltimos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rS As New ADODB.Recordset
Dim cnB As New clsPruebaContab

Private Sub cmbNCierres_Click()
    Label1 = "Últimos " + cmbNCierres + " Cierres"
    Set DGC.DataSource = cnB.GetRsCierres(CLng(cmbNCierres))
    AcomodarDG
End Sub

Private Sub cmdImprimir_Click()
    Dim PR As New tbrPrintRs.PrintRecorset
    
    Set PR.RsToPrint = cnB.GetRsCierres(CLng(cmbNCierres))
    Set PR.DatagridConWiths = DGC
    
    PR.LineasSeparadoras = True
    PR.Horizontal = True
    
    PR.Titulo = "Ultimos Cierres"
    
    PR.ImprimirRS
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmbNCierres = "10"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set DGC.DataSource = Nothing
    If rS.State = adStateOpen Then rS.Close
    Set rS = Nothing
    cnB.CN_CLOSE
    Set cnB = Nothing
End Sub

Private Sub AcomodarDG()
    DGC.Columns("ID").Width = 0
    'DGC.Columns("IdCierre").Caption = "IDC"
    'DGC.Columns("idc").Width = 500
    'DGC.Columns("idc").Alignment = dbgCenter
    DGC.Columns("IdCierre").Width = 500
    DGC.Columns("IdCierre").Alignment = dbgCenter
    DGC.Columns("sCierre").Width = 5700
    DGC.Columns("Fecha").Width = 900
    DGC.Columns("Fecha").Alignment = dbgCenter
    DGC.Columns("Hora").Width = 1000
    DGC.Columns("Hora").Alignment = dbgCenter
    DGC.Columns("Efvo").Width = 1000
    DGC.Columns("Efvo").Alignment = dbgCenter
    DGC.Columns("Efvo").NumberFormat = "$0.00"
    DGC.Columns("Diferencia").Width = 1000
    DGC.Columns("Diferencia").Alignment = dbgCenter
    DGC.Columns("Diferencia").NumberFormat = "$0.00"
    DGC.Columns("Variacion").Width = 1000
    DGC.Columns("Variacion").Alignment = dbgCenter
    DGC.Columns("Variacion").NumberFormat = "$0.00"
    
    DGC.RowHeight = 750
End Sub

Public Function AbrirDatos(MDB As String, PSW As String)
    cnB.ArchMDB = MDB
    cnB.PSW = PSW
    cnB.Conectar
    
    Me.Show 1
    
End Function

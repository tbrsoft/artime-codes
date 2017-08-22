VERSION 5.00
Begin VB.UserControl MouTextBox 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   2190
   ScaleWidth      =   4800
   ToolboxBitmap   =   "MouTextBox.ctx":0000
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3765
   End
End
Attribute VB_Name = "MouTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Event Declarations:
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Change() 'MappingInfo=Text1,Text1,-1,Change
'Default Property Values:
Const m_def_Largo = 1000
Const m_def_Entero = False

'Property Variables:
Dim m_Largo As Long
Dim m_Entero As Boolean

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    'primero que todo veo el largo (SOLO SI ESTA PINTADO LA DIF) y dejo es 8 que borra
    If Len(Text1) - Len(Text1.SelText) >= m_Largo And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
    Select Case KeyAscii
        
        Case 46, 44
            If m_Entero = True Then 'SI ES ENTERO no dejo que escriba comas
                KeyAscii = 0
            Else
                'chetada si esta pintada la parte que tiene coma permitir escribir la coma
                If CantidadCaracterEnCadena(",", Text1) - _
                    CantidadCaracterEnCadena(",", Text1.SelText) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = 44
                End If
            End If
        Case 45
            If m_Entero = True Then 'ENTERO ES enteros positivos
                KeyAscii = 0
            Else
                 'si esta todo pintado
                If Len(Text1) = Len(Text1.SelText) Then KeyAscii = 45: Exit Sub
                 
                 'para que no ponga menos en un lugar que no sea al principio
                If Len(Text1) > 0 Then KeyAscii = 0: Exit Sub
                
                 'normal si ya tiene un signo menos no permitir otro
                If CantidadCaracterEnCadena("-", Text1) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = 45
                End If
            End If
        
        Case 8, 13, 26, 27
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
        
End Sub

Private Function CantidadCaracterEnCadena(Caracter As String, Cadena As String) As Long
    Dim L As Long, T As Long
    
    T = 0
    L = Len(Cadena)
    For A = 1 To L
        If Mid(Cadena, A, 1) = Caracter Then T = T + 1
    Next A
    
    CantidadCaracterEnCadena = T
End Function

Private Sub UserControl_Resize()
    Text1.Top = 0
    Text1.Left = 0
    Text1.Width = UserControl.Width
    Text1.Height = UserControl.Height
End Sub


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Text1,Text1,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Devuelve o establece la alineación de un control CheckBox u OptionButton, o el texto de un control."
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    Text1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
    Text1.BackColor = PropBag.ReadProperty("BackColor", &HE0E0E0)
    Text1.Text = PropBag.ReadProperty("Text", "")
    Text1.Enabled = PropBag.ReadProperty("Enabled", Verdadero)
    Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
'    m_Largo = PropBag.ReadProperty("Largo", m_def_Largo)
'    m_Entero = PropBag.ReadProperty("Entero", m_def_Entero)
    m_Largo = PropBag.ReadProperty("Largo", m_def_Largo)
    m_Entero = PropBag.ReadProperty("Entero", m_def_Entero)
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", Text1.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("Text", Text1.Text, "")
    Call PropBag.WriteProperty("Enabled", Text1.Enabled, Verdadero)
    Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
'    Call PropBag.WriteProperty("Largo", m_Largo, m_def_Largo)
'    Call PropBag.WriteProperty("Entero", m_Entero, m_def_Entero)
    Call PropBag.WriteProperty("Largo", m_Largo, m_def_Largo)
    Call PropBag.WriteProperty("Entero", m_Entero, m_def_Entero)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Text1,Text1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub Text1_Change()
    RaiseEvent Change
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Devuelve o establece el texto contenido en el control."
Attribute Text.VB_UserMemId = 0
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Text1,Text1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Text1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Text1,Text1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get SelStart() As Long
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(NewStart As Long)
    Text1.SelStart = NewStart
End Property

Public Property Get SelLength() As Long
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(NewLength As Long)
    Text1.SelLength = NewLength
End Property

Public Function SetFocus()
    Text1.SetFocus
End Function

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
'    m_Largo = m_def_Largo
'    m_Entero = m_def_Entero
    m_Largo = m_def_Largo
    m_Entero = m_def_Entero
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,1000
Public Property Get Largo() As Long
    Largo = m_Largo
End Property

Public Property Let Largo(ByVal New_Largo As Long)
    m_Largo = New_Largo
    PropertyChanged "Largo"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get Entero() As Boolean
    Entero = m_Entero
End Property

Public Property Let Entero(ByVal New_Entero As Boolean)
    m_Entero = New_Entero
    PropertyChanged "Entero"
End Property


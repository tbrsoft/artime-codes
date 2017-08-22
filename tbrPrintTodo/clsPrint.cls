VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mFuenteNombre As String
Private mLineasSeparadoras As Boolean

Public Function ImprimirDATAG(DataG As Object, Titulo() As String, _
    TitColSep As String, Optional PiePag As String = "", _
    Optional ApartirDeCol As Long = 1, Optional Ensanchar As Single = 1, _
    Optional Pie2 As String = "", Optional Pie3 As String = "", _
    Optional Horizontal As Boolean = False, Optional RengXPagina As Long = 25) As Long
    'si ensanchar es < 0.65 es que debo imprimir tiquet
    'Tiene 5 titulos si o si! el tiquet no muestra el (3)
    
    '0 OK
    '1 esta vacio
    
    Dim A As Long, B As Long, MiY As Single, NPaginas As Long, EPagina As Long
    Dim tmP1 As String, tMp2 As String, MiX As Single
    Dim TitCol() As String, RxP As Long 'es Renglones por pagina
    Dim EsTiquet As Boolean, mHorizontal As Boolean
    Dim MaxX As Single, MaxY As Single, Ajustar As Single, SumAncho As Single
    Dim DataGD As DataGrid, RSdg As ADODB.Recordset, L As Long
    
    Set DataGD = DataG
    mHorizontal = Horizontal
    
    If DataGD.ApproxCount <= 0 Then
        ImprimirDATAG = 1
        Set DataGD = Nothing
        Exit Function
    End If
    
    Set RSdg = DataGD.DataSource
    
    Printer.Font.Name = mFuenteNombre
    If EsTiquet = False Then Printer.Font.Size = 12
    Printer.Font.Bold = True
        
    If Ensanchar <= 0.65 Then EsTiquet = True
    
    If EsTiquet Then
        MaxX = 3200: MaxY = 7500: RxP = 15
        Printer.FontSize = 8
    Else
        If mHorizontal = False Then
            MaxX = 9000: MaxY = 13000
            'solo se puede configurar rxP si es vertical
            RxP = RengXPagina
        Else
            MaxX = 13000: MaxY = 10000: RxP = 15
        End If
    End If
    
    TitCol = Split(TitColSep, "|")
     
     'otra vez si no se pierde
    If mHorizontal And EsTiquet = False Then
        Printer.Orientation = 2
    Else
        Printer.Orientation = 1
    End If
    
    NPaginas = RedondeoArriba(DataGD.ApproxCount / RxP)
    EPagina = (ApartirDeCol + RxP - 1) / RxP 'deberia ser entero siempre
        
    If RSdg.RecordCount - ApartirDeCol > RxP Then
        L = Printer.Orientation 'la pierde luego del endDoc
         'va a imprimir de atras para adelante
        ImprimirDATAG DataGD, Titulo, TitColSep, PiePag, ApartirDeCol + RxP, _
            Ensanchar, Pie2, Pie3, mHorizontal
            
        Printer.Orientation = L 'recupero la orientacion
    End If
    
    If EsTiquet = False Then Printer.FontSize = 12
    tPrint 400, 400, NoNuloS(Titulo(4)), , True, MaxX - 400, , mHorizontal
        
    If EsTiquet = False Then Printer.FontSize = 10
    'datos de mi empresa!!!!!!!!!!!!!!
    tPrint 400, 800, NoNuloS(Titulo(0))
    tPrint 400, 1200, NoNuloS(Titulo(1))
    tPrint 400, 1500, NoNuloS(Titulo(2))
    If EsTiquet = False Then
        tPrint MaxX, 1500, NoNuloS(Titulo(3)), True
    End If
    
    If EsTiquet = False Then
        tPrint MaxX, 400, "Fecha: " + FormatDateTime(Date, vbShortDate), True
    Else
        tPrint 400, 600, "Fecha: " + FormatDateTime(Date, vbShortDate)
    End If
    
    SumAncho = 0
    Dim Anchos() As Single, XX As Single
    For L = 0 To UBound(TitCol)
        ReDim Preserve Anchos(L)
        Anchos(L) = DataGD.Columns(L).Width * Ensanchar
        SumAncho = SumAncho + Anchos(L)
    Next L
    
    If SumAncho > MaxX Then 'se pasan -> ajusto al ancho
        Ajustar = MaxX / SumAncho
        For L = 0 To UBound(TitCol)
            Anchos(L) = Anchos(L) * Ajustar
        Next L
        SumAncho = MaxX - 150 'para saber hasta donde va el cuadrado!
    End If
    
    If EsTiquet = False Then Printer.FontSize = 9
    
    XX = 400: MiY = 2100
    For L = 0 To UBound(TitCol)
        If Anchos(L) > 0 Then
            If XX >= MaxX Then XX = MaxX
            tPrint XX + 100, MiY, CortarPr(TitCol(L), Anchos(L)), , , Anchos(L) - 30, , mHorizontal
            XX = XX + Anchos(L)
        End If
    Next L
    
    Printer.Font.Bold = False
    Printer.Font.Size = 8
    
    MiY = MiY + 100
    B = 1
    If RSdg.RecordCount > 0 Then
        RSdg.MoveFirst
        If ApartirDeCol > 1 Then RSdg.Move (ApartirDeCol - 1)
        Do While Not RSdg.EOF
            MiY = MiY + (230 * B)
            B = 1
            XX = 400 + Anchos(0) + 50
            If Anchos(0) > 0 Then tPrint 400, MiY, DataGD.Columns(0)
            For L = 1 To UBound(TitCol)
                If Anchos(L) > 0 Then
                    A = tPrint(XX + 100, MiY, DataGD.Columns(L), , , Anchos(L) - 50, , mHorizontal)
                    If A > B Then B = A
                    XX = XX + Anchos(L)
                    If XX >= MaxX Then XX = MaxX
                End If
            Next L
            RxP = RxP - 1
            If RxP = 0 Then RSdg.MoveLast
            RSdg.MoveNext
        Loop
    End If
    
    MiY = MiY + (230 * B - 1)
    'le pongo las lineas ahora que tengo que el alto y lo encuadro '--------------------
    If mLineasSeparadoras Then
        XX = 400 + Anchos(0) - 60
        For L = 1 To UBound(TitCol)
            If Anchos(L) > 0 Then
                PrintLINE XX, 2050, XX, MiY - 1700 + 2050, False
                XX = XX + Anchos(L)
                If L = 1 Then XX = XX + 100 'solo una vez
            End If
        Next L
        PrintRectangulo 330, 2050, MiY - 1700, SumAncho + 150
    End If
    '--------------------------------------------------------------------------------
    
    If NPaginas > 1 Then
        Printer.FontSize = 8
        tPrint 400, MaxY - 700, "Pagina " + CStr(EPagina) + " de " + CStr(NPaginas)
    End If
    
    If EsTiquet = False Then Printer.FontSize = 10
    tPrint MaxX, MaxY - 700, PiePag, True
    If Pie2 <> "" Then tPrint MaxX, MaxY - 900, Pie2, True
    If Pie3 <> "" Then tPrint MaxX, MaxY - 1100, Pie3, True
        
    PrintLINE MaxX - 3600, MaxY - 400, MaxX, MaxY - 400
    Printer.FontBold = True
    
    tPrint MaxX, MaxY - 300, "tbrSoft Desafios Digitales", True
    tPrint MaxX, MaxY, "CopyRight 2007 - info@tbrsoft.com", True
        
    Printer.EndDoc
    
    ImprimirDATAG = 0
    
    Set RSdg = Nothing
    Set DataGD = Nothing
End Function

Public Function ImprimirlvW(LvieW As Object, Titulo() As String, _
    TitColSep As String, Optional PiePag As String = "", _
    Optional ApartirDeCol As Long = 1, Optional Ensanchar As Single = 1, _
    Optional Pie2 As String = "", Optional Pie3 As String = "", _
    Optional Horizontal As Boolean = False, Optional RengXPagina As Long = 25, _
    Optional Ultimos As Long = 0) As Long
    'si ensanchar es < 0.65 es que debo imprimir tiquet
    'Tiene 5 titulos si o si! el tiquet no muestra el (3)
    
    '0 OK
    '1 esta vacio
    
    '------------------ ULTIMOS ----------------------
    '0 todo
    '-X Días atras (solo para cuando fecha sea la primer columna)
    ' X Movimientos atras
    '-------------------------------------------------
    
    
    Dim A As Long, B As Long, C As Long, MiY As Single, NPaginas As Long, EPagina As Long
    Dim tmP1 As String, tMp2 As String, MiX As Single
    Dim TitCol() As String, RxP As Long 'es Renglones por pagina
    Dim EsTiquet As Boolean, mHorizontal As Boolean
    Dim MaxX As Single, MaxY As Single, Ajustar As Single, SumAncho As Single
    Dim LvW As ListView, Hta As Long
    Dim Renglones As Long, J As Long, FechaRes As String 'depende de "Ultimo"
    
    Renglones = 0: FechaRes = ""
    Set LvW = LvieW
    mHorizontal = Horizontal
    
    If LvW.ListItems.Count = 0 Then
        ImprimirlvW = 1
        Set DataGD = Nothing
        Exit Function
    End If
    
    Printer.Font.Name = mFuenteNombre
    If EsTiquet = False Then Printer.Font.Size = 12
    Printer.Font.Bold = True
        
    If Ensanchar <= 0.65 Then EsTiquet = True
    
    If EsTiquet Then
        MaxX = 3200: MaxY = 7500: RxP = 15
        Printer.FontSize = 8
    Else
        If mHorizontal = False Then
            MaxX = 9000: MaxY = 13000: RxP = RengXPagina
        Else
            MaxX = 13000: MaxY = 10000: RxP = 15
        End If
    End If
    
    TitCol = Split(TitColSep, "|")
     
     'otra vez si no se pierde
    If mHorizontal And EsTiquet = False Then
        Printer.Orientation = 2
    Else
        Printer.Orientation = 1
    End If
    
    'defino cuanto es la cantidad de renglones a imprimir (hasta el ultimo) ----
    If Ultimos <> 0 Then
        If Ultimos > 0 Then
            If Ultimos > LvW.ListItems.Count Then
                Renglones = LvW.ListItems.Count
            Else
                Renglones = Ultimos
            End If
        Else 'veo cuantos renglones hay hasta esa fecha
            For J = 1 To LvW.ListItems.Count
                If Date + Ultimos - 1 > CDate(LvW.ListItems(J).Text) Then
                    Renglones = J
                    FechaRes = CStr(Date + Ultimos)
                    Exit For
                End If
            Next J
            
            'si la fecha es muy vieja pongo todo
            If Renglones = 0 Then Renglones = LvW.ListItems.Count
        End If
    Else
        'eligio imprimir todo
        Renglones = LvW.ListItems.Count
    End If
    
    ' agrego un renglon para detallar el SALDO AL "FECHA" -------
    ' (solo si sobran renglones)
    
    If Renglones < LvW.ListItems.Count Then
        Renglones = Renglones + 1
    End If
    ' -------------------------------------------------------------------------
        
    NPaginas = RedondeoArriba(Renglones / RxP)
    EPagina = (ApartirDeCol + RxP - 1) / RxP 'deberia ser entero siempre
        
    If Renglones - ApartirDeCol > RxP Then
        L = Printer.Orientation 'la pierde luego del endDoc
         'va a imprimir de atras para adelante
        ImprimirlvW LvW, Titulo, TitColSep, PiePag, ApartirDeCol + RxP, _
            Ensanchar, Pie2, Pie3, mHorizontal
            
        Printer.Orientation = L 'recupero la orientacion
        Hta = ApartirDeCol + RxP
    Else
        Hta = Renglones
    End If
    
    If EsTiquet = False Then Printer.FontSize = 12
    tPrint 400, 400, NoNuloS(Titulo(4)), , True, MaxX - 400, , mHorizontal
        
    If EsTiquet = False Then Printer.FontSize = 10
    
    'datos de mi empresa!!!!!!!!!!!!!!
    tPrint 400, 800, NoNuloS(Titulo(0))
    tPrint 400, 1200, NoNuloS(Titulo(1))
    tPrint 400, 1500, NoNuloS(Titulo(2))
    If EsTiquet = False Then
        tPrint MaxX, 1500, NoNuloS(Titulo(3)), True
    End If
    
    If EsTiquet = False Then
        tPrint MaxX, 400, "Fecha: " + FormatDateTime(Date, vbShortDate), True
    Else
        tPrint 400, 600, "Fecha: " + FormatDateTime(Date, vbShortDate)
    End If
    
    SumAncho = 0
    Dim Anchos() As Single, XX As Single
    
    For L = 0 To UBound(TitCol)
        ReDim Preserve Anchos(L)
        Anchos(L) = LvW.ColumnHeaders(L + 1).Width * Ensanchar
        SumAncho = SumAncho + Anchos(L)
    Next L
    
    If SumAncho > MaxX Then 'se pasan -> ajusto al ancho
        Ajustar = MaxX / SumAncho
        For L = 0 To UBound(TitCol)
            Anchos(L) = Anchos(L) * Ajustar
        Next L
        SumAncho = MaxX - 150 'para saber hasta donde va el cuadrado!
    End If
    
    If EsTiquet = False Then Printer.FontSize = 9
    
    XX = 400: MiY = 2100
    For L = 0 To UBound(TitCol)
        If Anchos(L) > 0 Then
            If XX >= MaxX Then XX = MaxX
            tPrint XX + 100, MiY, CortarPr(TitCol(L), Anchos(L)), , , Anchos(L) - 30, , mHorizontal
            XX = XX + Anchos(L)
        End If
    Next L
    
    Printer.Font.Bold = False
    Printer.Font.Size = 8
    
    MiY = MiY + 100
    B = 1
    For A = ApartirDeCol To Hta
        MiY = MiY + (230 * B)
        B = 1
        XX = 400 + Anchos(0) + 50
        '1ra columna ------------------------------------------------
        If Ultimos < 0 And A = Hta Then
            'si no imprime todo sino hasta cierta fecha
            'pongo esa fecha para empezar
            If Anchos(0) > 0 Then B = tPrint(400, MiY, FechaRes, _
                , , Anchos(0) - 50, , mHorizontal)
        Else
            If Anchos(0) > 0 Then B = tPrint(400, MiY, LvW.ListItems(A).Text, _
                , , Anchos(0) - 50, , mHorizontal)
        End If
        
        'Otras columnas ----------------------------------------------
        For L = 1 To UBound(TitCol)
            If Anchos(L) > 0 Then
                'si es la ultima fila deberia poner SALDO AL "FECHA"
                ' (suponiendo que tiene una columna de saldo (ultima))
                'SOLO CUANDO NO ESTA IMPRIMIENDO TODO
                'en la ¡Columna 2! tomo como detalle si hay problemas mas
                'adelante configurar
                'en la ¡Columna 1! tomo variacion y tiene que estar el saldo
                
                'imprimo Detalle y variacion ------------------------------------------------
                If A = Hta And Ultimos <> 0 And (L = 2 Or L = 1) Then
                    If FechaRes = "" And Ultimos > 0 Then 'es la que esta en la col 1
                        FechaRes = LvW.ListItems(Renglones).Text
                    End If
                    
                    If L = 2 Then
                        C = tPrint(XX + 100, MiY, "Saldo al " + FechaRes, , , _
                            Anchos(L) - 50, , mHorizontal)
                    Else 'l=1
                        C = tPrint(XX + 100, MiY, LvW.ListItems(A).SubItems(UBound(TitCol)), , , _
                            Anchos(L) - 50, , mHorizontal)
                    End If
                Else
                    C = tPrint(XX + 100, MiY, LvW.ListItems(A).SubItems(L), , , _
                        Anchos(L) - 50, , mHorizontal)
                End If
                '----------------------------------------------------------------
                
                If C > B Then B = C
                XX = XX + Anchos(L)
                If XX >= MaxX Then XX = MaxX
            End If
        Next L
        RxP = RxP - 1
        If RxP = 0 Then Exit For
    Next A
    
    MiY = MiY + (230 * B - 1)
    
    'le pongo las lineas ahora que tengo que el alto y lo encuadro '--------------------
    If mLineasSeparadoras Then
        XX = 400 + Anchos(0) - 60
        For L = 1 To UBound(TitCol)
            If Anchos(L) > 0 Then
                PrintLINE XX, 2050, XX, MiY - 1700 + 2050, False
                XX = XX + Anchos(L)
                If L = 1 Then XX = XX + 100 'solo una vez
            End If
        Next L
        PrintRectangulo 330, 2050, MiY - 1700, SumAncho + 150
    End If
    '--------------------------------------------------------------------------------
    
    If NPaginas > 1 Then
        Printer.FontSize = 8
        tPrint 400, MaxY - 700, "Pagina " + CStr(EPagina) + " de " + CStr(NPaginas)
    End If
    
    If EsTiquet = False Then Printer.FontSize = 10
    tPrint MaxX, MaxY - 700, PiePag, True
    If Pie2 <> "" Then tPrint MaxX, MaxY - 900, Pie2, True
    If Pie3 <> "" Then tPrint MaxX, MaxY - 1100, Pie3, True
        
    PrintLINE MaxX - 3600, MaxY - 400, MaxX, MaxY - 400
    Printer.FontBold = True
    
    tPrint MaxX, MaxY - 300, "tbrSoft Desafios Digitales", True
    tPrint MaxX, MaxY, "CopyRight 2007 - info@tbrsoft.com", True
        
    Printer.EndDoc
    
    ImprimirlvW = 0
    
    Set LvW = Nothing
End Function

Public Function CortarPr(StrP As String, Ancho As Single) As String
    Dim Resp As String, tmpS As String, Ancho2 As Single
    
    Ancho2 = Printer.TextWidth(StrP)
    If Ancho2 < Ancho Then
        Resp = StrP
    Else
        tmpS = StrP
        Do Until Ancho2 < Ancho - 150
            tmpS = Left(tmpS, Len(tmpS) - 1)
            Ancho2 = Printer.TextWidth(tmpS)
        Loop
        Resp = tmpS + "..."
    End If
    
    CortarPr = Resp
End Function

Public Function RedondeoArriba(Numero As Single) As Long
    Dim Resp As Long

    If Numero - Int(Numero) <> 0 Then 'tenia decimales
        Resp = Int(Numero) + 1 'redondeo arriba
    Else  'ya era un numero entero, lo dejo asi
        Resp = Numero
    End If
    
    RedondeoArriba = Resp
End Function

Public Sub PrintLINE(X As Single, Y As Single, Optional X2 As Single = 0, _
    Optional Y2 As Single = 0, Optional mStep As Boolean = False)
    
    Dim FinLine As Single, XX As Single, YY As Single
    
    FinLine = Printer.ScaleWidth * 0.9
    If X2 >= 0 Then
        XX = X2
    Else
        mStep = False
        XX = FinLine
    End If
    
    If Y2 >= 0 Then
        YY = Y2
    Else
        YY = Y
    End If
    
    If mStep Then
        XX = X + XX
        If Y <> YY Then
            YY = Y + YY
        End If
    End If
    
    Printer.DrawWidth = 8
    Printer.Line (X, Y)-(XX, YY)
End Sub

Private Sub PrintRectangulo(X As Single, Y As Single, Alto As Single, Optional Ancho As Single = 0)
    Dim FinLine As Single
    
    If Ancho = 0 Then
        FinLine = Printer.ScaleWidth * 0.9
    Else
        FinLine = Ancho
    End If
    
    Printer.DrawWidth = 12
    Printer.Line (X, Y)-Step(FinLine, Alto), , B
End Sub

Public Property Get FuenteNombre() As String
    FuenteNombre = mFuenteNombre
End Property

Public Property Let FuenteNombre(ByVal vNewValue As String)
    mFuenteNombre = vNewValue
End Property

Public Property Get LineasSeparadoras() As String
    LineasSeparadoras = mLineasSeparadoras
End Property

Public Property Let LineasSeparadoras(ByVal vNewValue As String)
    mLineasSeparadoras = vNewValue
End Property

Private Sub Class_Initialize()
    mFuenteNombre = "Arial"
    mLineasSeparadoras = False
End Sub

Private Function NoNuloN(J) As Single
    If IsNumeric(J) Then
        NoNuloN = J
    Else
        NoNuloN = 0
    End If
End Function

Private Function NoNuloS(S) As String
    If IsNull(S) Then
        NoNuloS = ""
    Else
        NoNuloS = S
    End If
End Function

Public Function tPrint(XX As Single, YY As Single, TXT As String, _
    Optional AlineadoDerecha As Boolean = False, _
    Optional Centrado As Boolean = False, _
    Optional Ancho As Single = 0, Optional SeparaPalabra As String = " ", _
    Optional Horizontal As Boolean = False, Optional Sigue As Boolean = True) As Long
    'da como resultado los renglones que uso
    
    Dim qXX As Long, Anch As Single, Tx2 As String, Tx3 As String
    Dim SP() As String, C As Long, Resp As Long, I As Long, MaxX As Single, MaxY As Single
        
    If Horizontal = False Then
        MaxX = 9000: MaxY = 16000
        'Printer.Orientation = 1
    Else
        MaxX = 13000: MaxY = 10000
        'Printer.Orientation = 2
    End If
    
    If Ancho = 0 Then
        If Horizontal = False Then
            Anch = 9000
        Else
            Anch = 13000
        End If
    Else
        Anch = Ancho
    End If
    
    Tx2 = "": Tx3 = "": Resp = 1
    
    'primero veo si hay saltos de pagina ('''''''''''''''''''''''''''''''''
    ''''' encontrar el cuadradito para hacer el salto de pagina!!!!!!!!
    If InStrRev(TXT, vbCrLf) <> 0 Then
        SP = Split(TXT, vbCrLf)
        For I = 0 To UBound(SP)
            YY = YY + (Printer.TextHeight(Right(SP(I), 1)) + 30) * I
            tPrint XX, YY, SP(I), AlineadoDerecha, Centrado, Ancho, SeparaPalabra
        Next I
        tPrint = I - 1
        Exit Function
    End If
    
    If Printer.TextWidth(TXT) > Anch Then
        'separa palabra por palabra
        Tx2 = TXT
        SP = Split(Tx2, SeparaPalabra)
        If UBound(SP) > 0 Then
           Do While Not Printer.TextWidth(Tx2) <= Anch
                C = UBound(SP)
                
                If C = 0 And Printer.TextWidth(Tx2) > Anch Then
                    Tx2 = CortarPr(Tx2, Anch - 10)
                Else
                    'primero elimino la última palabra a ver si asi si entra
                    If Tx3 <> "" Then
                        Tx3 = SP(C) + SeparaPalabra + Tx3
                    Else
                        Tx3 = SP(C)
                    End If
                    
                    ReDim Preserve SP(C - 1)
                    Tx2 = Join(SP, SeparaPalabra)
                End If
            Loop
            
            Resp = Resp + tPrint(XX, YY + Printer.TextHeight(Right(Tx2, 1)) + 30, Tx3, _
                AlineadoDerecha, Centrado, Anch, SeparaPalabra)
        Else 'no hay espacios corto a lo macho
        
        End If
    Else
        Tx2 = TXT
    End If
    
    If AlineadoDerecha = True Then
        qXX = XX - Printer.TextWidth(Tx2)
    Else
        If Centrado Then
            qXX = (XX + Anch - Printer.TextWidth(Tx2)) / 2
            If qXX < 50 Then qXX = 50
        Else
            qXX = XX
        End If
    End If
    
    If qXX <= 0 Then qXX = 200
    If YY <= 0 Then YY = 200
    
    Printer.CurrentX = qXX
    Printer.CurrentY = YY
    
    'si se pasa del margen derecho ??
    If (qXX + Printer.TextWidth(Tx2)) > Printer.Width - 1000 Then
        'ver hasta donde lleg
        Dim AA As Long, TTT As String
reINI:
        
        For AA = 1 To Len(Tx2)
            TTT = Mid(Tx2, 1, AA)
            If (qXX + Printer.TextWidth(TTT)) > Printer.Width - 1000 Then
                'salto de carro!
                'imprimo la primera parte y pàso abajo lo otro
                Printer.Print TTT
                Printer.CurrentX = qXX
                'Printer.CurrentY = Printer.CurrentY + Printer.TextHeight("aA")
                'txt ahora solo tendra lo primero
                Tx2 = Mid(Tx2, AA + 1, Len(Tx2) - AA)
                GoTo reINI
            End If
        Next AA
        
    End If
    
    tPrint = Resp
    Printer.Print Tx2;
    
    If Sigue = False Then Printer.EndDoc
End Function

Public Sub PrintCuadrado(BordeColor As Long, AnchoLinea As Long, _
                         X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
                         
    ' ------------- IMPRIMO CUADRADO -----------------------------------------------
    Dim BC As Long, AL As Long 'volores al inicio
    BC = Printer.ForeColor
    AL = Printer.DrawWidth
    
    Printer.ForeColor = BordeColor
    Printer.DrawWidth = AnchoLinea
    Printer.FillStyle = vbFSTransparent
    
    Printer.Line (X1, Y1)-(X2, Y2), , B
    
    Printer.ForeColor = BC
    Printer.DrawWidth = AL
    ' -----------------------------------------------------------------------------
End Sub

Public Sub SetFontPrinter(NameFont As String, SizeFont As Long, _
    oBold As Boolean, FontForeColor As Long)
    
    If NameFont <> "" Then Printer.Font.Name = NameFont
    If SizeFont <> 0 Then Printer.Font.Size = SizeFont
    Printer.Font.Bold = oBold
    If FontForeColor <> 0 Then Printer.ForeColor = FontForeColor
    
End Sub

Public Sub EndDocTP()
    Printer.EndDoc
End Sub


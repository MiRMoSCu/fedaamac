VERSION 5.00
Begin VB.Form Parametros 
   Caption         =   "Parametros"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtParGrupo 
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Text            =   " "
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label LblPar 
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Clave de Grupo para Estados de Cuenta"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PubNombre As String
Private Carpeta, PrvCita1, PrvCita2, PrvCita3, PrvSocio As String
Private flgsocio, totreg As Integer


Private Sub Form_Load()
    'IntRespuesta = MsgBox("Mi Primera Carpeta=" & frmMiPrimera.LblCarpeta, 0)
    Carpeta = frmMiPrimera.LblCarpeta
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM CITAS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cd.MoveFirst
    Do Until cd.EOF = True
        totreg = totreg + 1
        cd.MoveNext
        Loop
    If frmMiPrimera.Flg = 1 Then
       Label2.Caption = "INVERSION"
    Else
        Label2.Caption = "PRESTAMOS"
    End If
End Sub
Private Sub TxtParGrupo_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtParGrupo.Tag = MODE_OVERTYPE And TxtParGrupo.SelLength = 0 Then
        TxtParGrupo.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       If frmMiPrimera.Flg = 1 Then
            InverGrupo
        Else
            ParGrupo
        End If
       'SendKeys "{tab}"
    End If
    
    If flgsocio <> 1 Then
        If KeyAscii <> 13 Then
            flgsocio = 1
            TxtParGrupo = ""
        End If
    End If
End Sub
Private Sub ParGrupo()
Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim cp As New ADODB.Recordset 'Creamos el objeto Recordset.DMOVPR

   Dim strPath As String

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   cp.Open "SELECT * FROM DMOVPR ORDER BY GRUPO,SOCIO,FECHA,APREPAC DESC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    
    'Graba Encabezado del Estado de Cuenta
    Dim Word As Object
    Set Word = CreateObject("Word.Application")

    'Dim Word As New Word.Application

    'AGREGA  DOCUMENTO
    Dim LONGITUD As Single
    Word.Documents.Add

    cl.MoveFirst
    Selec_Grupo = TxtParGrupo
    
    Do Until cl.EOF = True
      If cl.Fields("GRUPO") = Selec_Grupo And cl.Fields("SALDOPRES") > 5 Then
        'MsgBox (cl.Fields("SOCIO") & "  " & cl.Fields("SALDOPRES"))
        PrvSocio = cl.Fields("SOCIO")
        BUSCA_SOCIO
        
        Word.Selection.TypeText "                   FONDO ECONOMICO DE AYUDA MUTUA, A.C" & vbCrLf
        Word.Selection.TypeText "                            ESTADO DE CUENTA" & vbCrLf
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
        Word.Selection.TypeText "Socio.-" & PrvSocio & ".-"
        LONGITUD = Len(PubNombre)
        LONGITUD = 48 - LONGITUD
        'MsgBox (cl.Fields("SOCIO") & "  " & PubNombre)
        Word.Selection.TypeText PubNombre & Space(LONGITUD)
        Word.Selection.TypeText "Fecha de Corte="
        Word.Selection.TypeText Format(cl.Fields("FECORTE"), "ddddd") & vbCrLf
        Word.Selection.TypeText "                     RESUMEN DE PRESTAMOS"
        Word.Selection.TypeText Space(19) & "Tasa de Interés      "
        Word.Selection.TypeText Format(cl.Fields("TASAPRES") / 100, "Percent") & vbCrLf
        
        Word.Selection.TypeText "Saldo Inicial      "
        Importe = Format(cl.Fields("PRES_INI"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("PRES_INI"), "Currency") & " |"
        
        Word.Selection.TypeText "Préstamos      "
        Importe = Format(cl.Fields("PRESTAMOS"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("PRESTAMOS"), "Currency") & " |"
        
        Word.Selection.TypeText "Fecha Prestamo="
        Word.Selection.TypeText Format(cl.Fields("FECPRES"), "ddddd") & vbCrLf
        
        Word.Selection.TypeText "Saldo Actual       "
        Importe = Format(cl.Fields("SALDOPRES"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("SALDOPRES"), "Currency") & " |"
        
        Word.Selection.TypeText "Pagos          "
        Importe = Format(cl.Fields("PAGOS"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("PAGOS"), "Currency") & " |"
        
        Word.Selection.TypeText "Vencimiento=   "
        Word.Selection.TypeText Format(cl.Fields("FECVENC"), "ddddd") & vbCrLf
        
        Word.Selection.TypeText "Pago Mínimo        "
        Importe = Format(cl.Fields("PAGOMIN"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("PAGOMIN"), "Currency") & " |"
        
        Word.Selection.TypeText "Pago Total     "
        Dim Pagotot As Single
        Pagotot = (cl.Fields("SALDOPRES") * cl.Fields("TASAPRES") / 100) + cl.Fields("SALDOPRES")
        Importe = Format(Pagotot, "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(Pagotot, "Currency") & " |"
        
        Word.Selection.TypeText "Ints Pagados  "
        Importe = Format(cl.Fields("INTPAGADO"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("INTPAGADO"), "Currency") & vbCrLf
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
        Word.Selection.TypeText "FECHA          DESCRIPCION          PAGOS        "
        Word.Selection.TypeText " PRESTAMOS       SALDO" & vbCrLf
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
        
        cp.MoveFirst
      Do Until cp.EOF = True
        x = x + 0.1
        Me.CurrentX = x
        Me.CurrentY = 6
        Me.Print ">"
        If cp.Fields("SOCIO") = PrvSocio Then
            Word.Selection.TypeText Format(cp.Fields("FECHA"), "ddddd") & " "
        
            DESCRIP = cp.Fields("DESCRIP")
            LONGITUD = Len(DESCRIP)
            LONGITUD = 25 - LONGITUD
            Word.Selection.TypeText DESCRIP & Space(LONGITUD)
        
            Importe = Format(cp.Fields("IMPORTE"), "Currency")
            LONGITUD = Len(Importe)
            LONGITUD = 11 - LONGITUD
            If cp.Fields("APREPAC") = "P" Then
                  '*Abonos
                Word.Selection.TypeText Space(LONGITUD) & Format(cp.Fields("IMPORTE"), "Currency") & Space(16)
                sdoActual = sdoActual - cp.Fields("IMPORTE")
            Else
                '      *Cargos
                Word.Selection.TypeText Space(13)
                Word.Selection.TypeText Space(LONGITUD) & Format(cp.Fields("IMPORTE"), "Currency") & "   "
                sdoActual = sdoActual + cp.Fields("IMPORTE")
            End If
            Importe = Format(sdoActual, "Currency")
            LONGITUD = Len(Importe)
            LONGITUD = 11 - LONGITUD
            Word.Selection.TypeText Space(LONGITUD) & Format(sdoActual, "Currency") & " "
            If cp.Fields("CTABCO") <> "" Then
                Word.Selection.TypeText cp.Fields("CTABCO") & " "
            Else
                Word.Selection.TypeText "    "
            End If
            Word.Selection.TypeText cp.Fields("REFERENC") & vbCrLf
        End If
        cp.MoveNext

        Loop
    
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
        'Word.Selection.TypeText vbPageBreaks

          'InsertBreak Type wdPageBreak


        Busca_Cita
        Word.Selection.TypeText Space(20) & PrvCita1 & vbCrLf
        Word.Selection.TypeText Space(20) & PrvCita2 & vbCrLf
        Word.Selection.TypeText Space(20) & PrvCita3 & vbCrLf
        
        Word.Selection.TypeText vbCrLf
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
        Word.Selection.TypeText vbCrLf
        Word.Selection.TypeText vbCrLf

        sdoActual = 0
    End If
    cl.MoveNext
    Loop
    'MsgBox (PrvSocio & ".-" & PubNombre)
        'AGREGA PARRAFO
        Word.Selection.TypeParagraph
    
    
    'SELECCIONA TEXTO
    Word.Selection.WholeStory
    Word.Selection.Font.Size = 8
    
    
    ' VISIBLE
    Word.Visible = True

    Set Word = Nothing
    MsgBox ("Se generó Estados de Cuenta de PRESTAMOS en WORD")
    flgsocio = 0
    Unload Me
    Static lfrmCount As Long
    Dim frmD As Parametros
    lfrmCount = lfrmCount + 1
    Set frmD = New Parametros
    frmD.Caption = "Parametros"
    
    frmD.Show
End Sub
Private Sub InverGrupo()
Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim cp As New ADODB.Recordset 'Creamos el objeto Recordset.DMOVPR

   Dim strPath As String

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   cp.Open "SELECT * FROM DMOVIN ORDER BY GRUPO,SOCIO,FECHA,APREPAC DESC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    
    'Graba Encabezado del Estado de Cuenta
    Dim Word As Object
    Set Word = CreateObject("Word.Application")

    'Dim Word As New Word.Application

    'AGREGA  DOCUMENTO
    Dim LONGITUD As Single
    Word.Documents.Add

    cl.MoveFirst
    Selec_Grupo = TxtParGrupo
    
    Do Until cl.EOF = True
      If cl.Fields("GRUPO") = Selec_Grupo And cl.Fields("INTGANADO") > 0 Then
        'MsgBox (cl.Fields("SOCIO") & "  " & cl.Fields("SALDOPRES"))
        PrvSocio = cl.Fields("SOCIO")
        BUSCA_SOCIO
        Word.Selection.TypeText "                   FONDO ECONOMICO DE AYUDA MUTUA, A.C" & vbCrLf
        Word.Selection.TypeText "                            ESTADO DE CUENTA" & vbCrLf
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
        Word.Selection.TypeText "Socio.-" & PrvSocio & ".-"
        LONGITUD = Len(PubNombre)
        LONGITUD = 48 - LONGITUD
        Word.Selection.TypeText PubNombre & Space(LONGITUD)
        Word.Selection.TypeText "Fecha de Corte="
        Word.Selection.TypeText Format(cl.Fields("FECORTE"), "ddddd") & vbCrLf
        Word.Selection.TypeText "                     RESUMEN DE INVERSION"
        Word.Selection.TypeText Space(19) & "Tasa de Interés       "
        Word.Selection.TypeText Format(cl.Fields("INTGANADO") / cl.Fields("PROM_INV"), "Percent") & vbCrLf
        
        Word.Selection.TypeText "Saldo Inicial      "
        Importe = Format(cl.Fields("INV_INI"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("INV_INI"), "Currency") & " |"
        
        Word.Selection.TypeText "Aportaciones   "
        Importe = Format(cl.Fields("APORTA"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("APORTA"), "Currency") & " |"
        
        Word.Selection.TypeText "Fecha Apertura= "
        Word.Selection.TypeText Format(cl.Fields("FECAPER"), "ddddd") & vbCrLf
        
        Word.Selection.TypeText "Saldo Actual       "
        Importe = Format(cl.Fields("SALDO"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("SALDO"), "Currency") & " |"
        
        Word.Selection.TypeText "Retiros        "
        Importe = Format(cl.Fields("RETIROS"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("RETIROS"), "Currency") & " |"
        
        Word.Selection.TypeText "Fecha de Nac.   "
        Word.Selection.TypeText Format(cl.Fields("FECNAC"), "ddddd") & vbCrLf
        
        Word.Selection.TypeText "Int. Devengados    "
        Importe = Format(cl.Fields("INTGANADO"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("INTGANADO"), "Currency") & " |"
                       
        Word.Selection.TypeText "Saldo Promedio "
        Importe = Format(cl.Fields("PROM_INV"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(cl.Fields("PROM_INV"), "Currency") & " |"
        
        Word.Selection.TypeText "Prom.Aportación"
        Dim Pagotot As Single
        If Month(cl.Fields("FECORTE")) > 10 Then
            varNumMes = Month(cl.Fields("FECORTE")) - 10
          Else
            varNumMes = Month(cl.Fields("FECORTE")) + 2
          End If
        Pagotot = Format(cl.Fields("APORTA") / varNumMes, "Currency")
        Importe = Format(Pagotot, "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(Pagotot, "Currency") & vbCrLf
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
        Word.Selection.TypeText "FECHA          DESCRIPCION          RETIROS      "
        Word.Selection.TypeText "APORTACIONES    SALDO" & vbCrLf
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf

        cp.MoveFirst
 Do Until cp.EOF = True
        x = x + 0.1
        Me.CurrentX = x
        Me.CurrentY = 6
        Me.Print ">"
   If cp.Fields("SOCIO") = PrvSocio Then
        Word.Selection.TypeText Format(cp.Fields("FECHA"), "ddddd") & " "
        
        DESCRIP = cp.Fields("DESCRIP")
        LONGITUD = Len(DESCRIP)
        LONGITUD = 25 - LONGITUD
        Word.Selection.TypeText DESCRIP & Space(LONGITUD)
        
        Importe = Format(cp.Fields("IMPORTE"), "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        If cp.Fields("APREPAC") = "R" Then
                  '*Abonos
            Word.Selection.TypeText Space(LONGITUD) & Format(cp.Fields("IMPORTE"), "Currency") & Space(16)
                sdoActual = sdoActual - cp.Fields("IMPORTE")
        Else
            '      *Cargos
            Word.Selection.TypeText Space(13)
            Word.Selection.TypeText Space(LONGITUD) & Format(cp.Fields("IMPORTE"), "Currency") & "   "
            sdoActual = sdoActual + cp.Fields("IMPORTE")
        End If
        Importe = Format(sdoActual, "Currency")
        LONGITUD = Len(Importe)
        LONGITUD = 11 - LONGITUD
        Word.Selection.TypeText Space(LONGITUD) & Format(sdoActual, "Currency") & " "
        If cp.Fields("CTABCO") <> "" Then
            Word.Selection.TypeText cp.Fields("CTABCO") & " "
        Else
            Word.Selection.TypeText "    "
        End If
        Word.Selection.TypeText cp.Fields("REFERENC") & vbCrLf
   End If
   cp.MoveNext

 Loop
   Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
   'Word.Selection.TypeText vbPageBreaks

   'InsertBreak Type wdPageBreak


   Busca_Cita
   Word.Selection.TypeText Space(20) & PrvCita1 & vbCrLf
   Word.Selection.TypeText Space(20) & PrvCita2 & vbCrLf
   Word.Selection.TypeText Space(20) & PrvCita3 & vbCrLf
   
   Word.Selection.TypeText vbCrLf
   Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
   Word.Selection.TypeText vbCrLf
   Word.Selection.TypeText vbCrLf

    sdoActual = 0
    End If
    cl.MoveNext
    Loop
    'AGREGA PARRAFO
        Word.Selection.TypeParagraph
    
    
    'SELECCIONA TEXTO
    Word.Selection.WholeStory
    Word.Selection.Font.Size = 8
    
    
    ' VISIBLE
    Word.Visible = True

    Set Word = Nothing
    MsgBox ("Se generó Estados de Cuenta  de INVERSION en WORD")
    flgsocio = 0
    Unload Me
    Static lfrmCount As Long
    Dim frmD As Parametros
    lfrmCount = lfrmCount + 1
    Set frmD = New Parametros
    frmD.Caption = "Parametros"
    
    frmD.Show
End Sub
Private Sub BUSCA_SOCIO()
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
    cl.MoveFirst
    'IntRespuesta = MsgBox("BUSCA_SOCIO=" & PrvSocio, 0)

    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = PrvSocio Then
            

            If cl.Fields("FECVENC") <> "" Then
                PrvMeses = (cl.Fields("FECVENC") - cl.Fields("FECPRES")) / 30.4
            End If
            PubNombre = cl.Fields("NOMBRE")
            'IntRespuesta = MsgBox("BUSCA_SOCIO=" & PrvSocio & ".-" & PubNombre, 0)
            PrvFecorte = cl.Fields("FECORTE")
            PrvFecPres = cl.Fields("FECPRES")
            PrvFecVenc = cl.Fields("FECVENC")
            PrvPagoMin = cl.Fields("PAGOMIN")
            PrvGrupo = cl.Fields("GRUPO")
            PrvInvini = cl.Fields("INV_INI")
            PrvSaldoInicial = cl.Fields("PRES_INI")
            PrvSaldo = cl.Fields("SALDO")
            prvPrestamos = cl.Fields("PRESTAMOS")
            PrvAporta = cl.Fields("APORTA")
            PrvRetiros = cl.Fields("RETIROS")
            prvIntPagado = cl.Fields("INTPAGADO")
            prvComision = cl.Fields("COMISION")
            PrvPagos = cl.Fields("PAGOS")
            prvIntGanado = cl.Fields("INTGANADO")
            s_saldopres = cl.Fields("SALDOPRES")
            s_tasapres = cl.Fields("TASAPRES")
            PrvSeguro = cl.Fields("CTASEGURO")
            prvPromotor = cl.Fields("PROMOTOR")
            prvPromedio = cl.Fields("PROM_INV")
            Exit Do
        End If
        cl.MoveNext
        Loop
End Sub
Private Sub Busca_Cita()
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM CITAS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cd.MoveFirst
    
    Do Until cd.EOF = True
        Randomize
        'varita = varita - 0.1
        Aleatorio = CLng((1 - totreg) * Rnd + totreg)
        'If Aleatorio < 0 Then
        '    Aleatorio = 1
        '    varita = 1
        'End If
        If Aleatorio = numreg Then
            'LblNreg = numreg
            'LblCita1.Caption = cd.Fields("CITA1")
            PrvCita1 = cd.Fields("CITA1")
            If cd.Fields("CITA2") > "" Then
                'LblCita2.Caption = cd.Fields("CITA2")
                PrvCita2 = cd.Fields("CITA2")
            Else
                PrvCita2 = ""
            End If
            If cd.Fields("CITA3") > "" Then
                'LblCita3.Caption = cd.Fields("CITA3")
                PrvCita3 = cd.Fields("CITA3")
            Else
                PrvCita3 = ""
            End If
            Exit Do
        End If
        numreg = numreg + 1
        cd.MoveNext
        Loop
        numreg = 1
cd.Close
End Sub




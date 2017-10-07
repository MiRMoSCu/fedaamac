VERSION 5.00
Begin VB.Form frmMENUSYS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Aplicación FEDAMAC"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   12210
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture4 
      Height          =   2655
      Left            =   3960
      Picture         =   "frmMENUSYS.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   3795
      TabIndex        =   32
      Top             =   4080
      Width           =   3855
   End
   Begin VB.PictureBox Picture3 
      Height          =   2415
      Left            =   7920
      Picture         =   "frmMENUSYS.frx":56A0
      ScaleHeight     =   2355
      ScaleWidth      =   3555
      TabIndex        =   31
      Top             =   4200
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   7920
      Picture         =   "frmMENUSYS.frx":9309
      ScaleHeight     =   1995
      ScaleWidth      =   3555
      TabIndex        =   26
      Top             =   2040
      Width           =   3615
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   2895
         TabIndex        =   30
         Top             =   2040
         Width           =   2895
      End
   End
   Begin VB.PictureBox PicSocio 
      Height          =   1335
      Left            =   6240
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   25
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox TxtPagoMin 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtTasa 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   5
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtMeses 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtCapital 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton CmdCalculo 
      Caption         =   "Cálculo de Pago Mínimo"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000D&
      Caption         =   "MAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2520
      TabIndex        =   29
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000D&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2160
      TabIndex        =   28
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "FED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1560
      TabIndex        =   27
      Top             =   240
      Width           =   735
   End
   Begin VB.Label LblcBanco 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre del Administrador"
      Height          =   375
      Left            =   1080
      TabIndex        =   23
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base de Datos"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   855
   End
   Begin VB.Label LblnBanco 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label LbldBanco 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LblTotIntereses 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Total de Intereses"
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label LblNreg 
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label LblCita3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label LblCita2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Label LblCita1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   2640
      Width           =   4815
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "MENSUAL"
      Height          =   255
      Left            =   7920
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "%"
      Height          =   255
      Left            =   10200
      TabIndex        =   12
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Pago Mínimo"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "TASA DE INT."
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "MESES"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "CAPITAL"
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   135
      Left            =   6840
      TabIndex        =   6
      Top             =   2760
      Width           =   15
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   15
      Left            =   2280
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblFedamac 
      AutoSize        =   -1  'True
      Caption         =   "FEDAMAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   885
   End
   Begin VB.Menu mnuMovimientos 
      Caption         =   "&Movimientos"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnucaptura 
         Caption         =   "&Captura"
      End
      Begin VB.Menu MnuReembolso 
         Caption         =   "&Captura Reembolso Automático"
      End
      Begin VB.Menu mnurelmov 
         Caption         =   "&Relación de Movimientos"
      End
      Begin VB.Menu MnuAgenda 
         Caption         =   "&Agenda"
      End
      Begin VB.Menu mnuArchivoSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuSocios 
      Caption         =   "&Socios"
      Begin VB.Menu mnuConSocios 
         Caption         =   "&Consulta de Socios"
      End
      Begin VB.Menu mnuListaNombres 
         Caption         =   "&BALANCE"
      End
      Begin VB.Menu mnuEctaGrupo 
         Caption         =   "&Lista Socios en FLEX GRID"
      End
      Begin VB.Menu MnuListaPrestamos 
         Caption         =   "&Socios Morosos"
      End
      Begin VB.Menu MnuListaPrestamos1 
         Caption         =   "&Prestamos Otorgados"
      End
   End
   Begin VB.Menu mnuImprimir 
      Caption         =   "&Imprimir"
      Begin VB.Menu MnuEdoCta 
         Caption         =   "&Estados de Cuenta de Préstamos"
      End
      Begin VB.Menu mnuEctaInv 
         Caption         =   "&Estados de Cuenta de Inversión"
      End
      Begin VB.Menu MnuRMAY 
         Caption         =   "&Socios Mayores"
      End
      Begin VB.Menu MnuRelSocios 
         Caption         =   "&Relación de Socios"
      End
      Begin VB.Menu MnuTotGrupo 
         Caption         =   "Total por Grupos"
      End
      Begin VB.Menu RelGrupos 
         Caption         =   "Relación de Grupos"
      End
      Begin VB.Menu MnuRpre 
         Caption         =   "Relación de Préstamos"
      End
      Begin VB.Menu Desglose50 
         Caption         =   "&Desglose Cta 50"
      End
      Begin VB.Menu EctaPreGrupo 
         Caption         =   "&Estados de Cuenta de Préstamos por Grupos"
      End
      Begin VB.Menu EctaInGrupo 
         Caption         =   "&Estados de Cuenta de Inversión por Grupos"
      End
   End
   Begin VB.Menu mnuCapitalizar 
      Caption         =   "&Inicialización de Archivos"
      Begin VB.Menu mnuAnual 
         Caption         =   "&Capitalización Anual por días"
      End
      Begin VB.Menu MnuInicioEjercicio 
         Caption         =   "&Inicio de Ejercicio"
      End
      Begin VB.Menu InicioCaptura 
         Caption         =   "&Inicio de Captura de Cuentahabiente"
      End
      Begin VB.Menu MnuDBGNL 
         Caption         =   "&Captura DBGNL"
      End
      Begin VB.Menu MnuDBJBS 
         Caption         =   "&CapturaDBJBS"
      End
      Begin VB.Menu MnuCapDBVCL 
         Caption         =   "&CapturaDBVCL"
      End
      Begin VB.Menu MnuCapDBGMF 
         Caption         =   "&CapturaDBGMF"
      End
      Begin VB.Menu MnuCapDBBMX 
         Caption         =   "&CapturaDBBMX"
      End
      Begin VB.Menu MnuCapDBLLB 
         Caption         =   "&CapturaDBLLB"
      End
   End
   Begin VB.Menu MnuConsultaCH 
      Caption         =   "&Consulta CuentaHabientes"
      Begin VB.Menu MnuConsultaDBGNL 
         Caption         =   "&DBGNL"
      End
      Begin VB.Menu MnuConsultaDBJBS 
         Caption         =   "&DBJBS"
      End
      Begin VB.Menu MnuConsultaDBVCL 
         Caption         =   "&DBVCL"
      End
      Begin VB.Menu MnuConsultaDBBMF 
         Caption         =   "&DBGMF"
      End
      Begin VB.Menu MnuConsultaDBBMX 
         Caption         =   "&DBBMX"
      End
      Begin VB.Menu MnuConsultaDBLLB 
         Caption         =   "&DBLLB"
      End
   End
End
Attribute VB_Name = "frmMENUSYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declaramos los objetos
'Public MSWord As New Word.Application

Public Carpeta As String
Public Documento As Object
Private Saldo_DBF, SaldoPres_DBF, IntGanado_DBF, Comision_DBF As Single

Private s_numreg, cNumes, IdNum, PorcienGrupo As Integer
Private cGrupo, cSocio As String
Private cCveMov, cTipo, cAPrePac, cDescrip, cCtaBco, cReferenc, cTasa As String
Private dBanco, cBanco, cBco As String
Private PrvSeguro, prvPromotor As String
Private caprestamo, cImporte, TotRetiros As Double
Private PrvMeses As Integer
Private PrvFecorte, PrvFecVenc, PrvFecPres, cFecha As Date
Private PrvSocio, PubNombre, PrvGrupo, s_tipo As String
Private f_final As Date
Private totreg, TotAbonos, TotCargos As Single
Private numreg As Single
Private Aleatorio As Long
Private s_saldopres, PrvPagoMin, s_tasapres As Single
Private s_intganado, PrvInvini As Single
Private PrvSaldoInicial, prvPrestamos, PrvSaldo As Single
Private PrvAporta, PrvRetiros, prvIntPagado, prvComision, PrvPagos, prvIntGanado As Single
Private PrvCita1, PrvCita2, PrvCita3 As String
Private prvInversion, prvPromedio As Single
Private tcomision, tinversion, tintganado, tintpagado, tpromedio As Single
Private TotPorciento, Porciento, TotReembolso As Single


Private Sub CmdCalculo_Click()

    Dim sp_importe, sp_plazo, sp_tasa, sp_pagomin As Single
    
    sp_importe = TxtCapital
    TxtCapital = Format(sp_importe, "###,###,##0.00")

    sp_plazo = TxtMeses
    sp_tasa = TxtTasa
    'IntRespuesta = MsgBox(sp_importe, vbOKCancel)

    If sp_importe <> "" Then
        sp_pagomin = (sp_importe / (((1 - 1 * ((1 + (sp_tasa / 100)) ^ -sp_plazo))) / (sp_tasa / 100)))
        LblTotIntereses = Format(sp_pagomin * sp_plazo - sp_importe, "$###,###,##0.00")
    Else
        IntRespuesta = MsgBox("FAVOR DE INSERTAR DATOS", vbOKCancel)
    End If
    TxtCapital = Format(sp_importe, "$###,###,##0.00")

    TxtPagoMin = Format(sp_pagomin, "Currency")
End Sub

Private Sub lblpsw_Click()
If txtpsw.Text = "txtpsw" Then
    Beep
    lblpsw.Caption = "Contraseña correcta"
    ImgCaptura.Picture = LoadPicture("\Archivos de Programa\Microsoft Visual Studio\Common\Graphics\Bitmaps\Assorted\Happy.bmp")
    MsgBox ("Contraseña correcta")
Else
    lblpsw.Caption = "Contraseña incorrecta"
    txtpsw.Text = ""
    ImgCaptura.Picture = LoadPicture("\Archivos de Programa\Microsoft Visual Studio\Common\Graphics\Bitmaps\Assorted\INTL_NO.bmp")
    IntRespuesta = MsgBox("Contraseña INCORRECTA", vbOKCancel)
    
End If
End Sub



Private Sub Desglose50_Click()
 Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim cp As New ADODB.Recordset 'Creamos el objeto Recordset.DMOVIN

   Dim strPath As String

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   cp.Open "SELECT * FROM SICMOV ORDER BY SOCIO,APREPAC,CVEMOV,FECHA,IMPORTE", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cl.MoveFirst
    PrvSocio = "50"
    BUSCA_SOCIO
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = PrvSocio Then
            Exit Do
        End If
        cl.MoveNext
        Loop
    

    'Graba Encabezado del Estado de Cuenta
    Dim Word As Object
    Set Word = CreateObject("Word.Application")

    'Dim Word As New Word.Application

    'AGREGA  DOCUMENTO
    Dim LONGITUD As Single
    Word.Documents.Add
        Word.Selection.TypeText "                   FONDO ECONOMICO DE AYUDA MUTUA, A.C" & vbCrLf
        Word.Selection.TypeText "                         DESGLOSE DE MOVIMIENTOS" & vbCrLf
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
        Word.Selection.TypeText "Socio.-" & PrvSocio & ".-"
        LONGITUD = Len(PubNombre)
        LONGITUD = 48 - LONGITUD
        Word.Selection.TypeText PubNombre & Space(LONGITUD)
        Word.Selection.TypeText " Fecha de Corte="
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
        cCveMov = "00"
 Do Until cp.EOF = True
   If cp.Fields("SOCIO") = PrvSocio And cp.Fields("IMPORTE") > 0 Then
        If cp.Fields("CVEMOV") <> cCveMov Then
         If TotRetiros > 0 Then
            Word.Selection.TypeText "                 SUB-TOTAL          "
            Importe = Format(TotRetiros, "Currency")
            LONGITUD = Len(Importe)
            LONGITUD = 11 - LONGITUD
            Word.Selection.TypeText Space(LONGITUD) & Format(TotRetiros, "Currency") & Space(16)
            Word.Selection.TypeText "" & vbCrLf
         End If
            Word.Selection.TypeText "" & vbCrLf

            TotRetiros = 0
            cCveMov = cp.Fields("CVEMOV")
        End If
        
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
                TotRetiros = TotRetiros + cp.Fields("IMPORTE")
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
    Word.Selection.TypeText "                 SUB-TOTAL          "
    Importe = Format(TotRetiros, "Currency")
    LONGITUD = Len(Importe)
    LONGITUD = 11 - LONGITUD
    Word.Selection.TypeText Space(LONGITUD) & Format(TotRetiros, "Currency") & Space(16) & vbCrLf
           
   Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
   'Word.Selection.TypeText vbPageBreaks

   'InsertBreak Type wdPageBreak


   Busca_Cita
   Word.Selection.TypeText Space(20) & PrvCita1 & vbCrLf
   Word.Selection.TypeText Space(20) & PrvCita2 & vbCrLf
   Word.Selection.TypeText Space(20) & PrvCita3 & vbCrLf
  
        'AGREGA PARRAFO
        Word.Selection.TypeParagraph
    
    
    'SELECCIONA TEXTO
    Word.Selection.WholeStory
    Word.Selection.Font.Size = 8
    
    
    ' VISIBLE
    Word.Visible = True

    Set Word = Nothing
    
 
 
   
      'IntRespuesta = MsgBox("MODO DE PRUEBA -NO DISPONIBLE-", 0)

   'IntRespuesta = MsgBox("Se generó Estado de Cuenta de Préstamos: ECTAPRESTAMO", 0)

    'Static lfrmCount As Long
    'Dim frmD As frmMENUSYS
    'lfrmCount = lfrmCount + 1
    'Set frmD = New frmMENUSYS
    'frmD.Caption = "frmMENUSYS"
    
    'frmD.Show

End Sub

Private Sub EctaInGrupo_Click()
    frmMiPrimera.Flg = 1
    Static lfrmCount As Long
    Dim frmD As Parametros
    lfrmCount = lfrmCount + 1
    Set frmD = New Parametros
    frmD.Caption = "Parametros"
    
    frmD.Show
End Sub

Private Sub EctaPreGrupo_Click()
    Static lfrmCount As Long
    Dim frmD As Parametros
    lfrmCount = lfrmCount + 1
    Set frmD = New Parametros
    frmD.Caption = "Parametros"
    
    frmD.Show
End Sub
   

Private Sub Form_Load()
    Carpeta = frmMiPrimera.LblCarpeta
    lblFedamac = frmMiPrimera.LblEmpresa
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.CITAS
   Dim cs As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM CITAS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    Do Until cd.EOF = True
        totreg = totreg + 1
        cd.MoveNext
        Loop
            Randomize

'    IntRespuesta = MsgBox("Mi Primera Carpeta=" & frmMiPrimera.LblCarpeta, 0)
 '   IntRespuesta = MsgBox("Carpeta=" & Carpeta, 0)
    PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Fedamac.jpg")

    dBanco = frmMiPrimera.txtpsw
    LbldBanco = dBanco
    If dBanco = "F3D4M4C" Then
        LblnBanco = "MARTHA PATRICIA LOPEZ BAEZA"
        LblcBanco = SISFED
    End If
    If dBanco = "DBGNL" Then
        LblnBanco = "JULIO NIETO MATA"
    End If
    If dBanco = "DBVCL" Then
        LblnBanco = "GABRIELA CARRILLO LOPEZ"
    End If
    If dBanco = "DBJBS" Then
        LblnBanco = "JESUS BAUTISTA SERNA"
    End If
    If dBanco = "DBGMF" Then
        LblnBanco = "GERARDO MONTES FUENTES"
    End If
    If dBanco = "DBBMX" Then
        LblnBanco = "LUIS LOPEZ BAEZA"
    End If

   'Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb" '''

'Randomize
'MsgBox ("LblCarpeta.-" & LblCarpeta)
'Dim nAleatorio As Single
'Dim v_entero As Long

   'Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
  
   'cd.Open "SELECT * FROM CITAS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
    cd.MoveFirst
   'ntRespuesta = MsgBox("CLng" & "=Aleatorio" & "=" & Aleatorio & "=" & Rnd & "=" & v_entero & "=" & totreg, 0)

Busca_Cita
   cs.Open "SELECT * FROM SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
     
    Do Until cs.EOF = True
        If Day(cs.Fields("FECNAC")) = Day(Date) And Month(cs.Fields("FECNAC")) = Month(Date) Then
            MsgBox ("¡¡¡FELICIDADES!!! Hoy es cumpleaños de: " & cs.Fields("NOMBRE"))
        End If
                If Day(cs.Fields("FECNAC")) - 1 = Day(Date) And Month(cs.Fields("FECNAC")) = Month(Date) Then
            MsgBox (cs.Fields("FECNAC") & " Es cumpleaños de: " & cs.Fields("NOMBRE"))
        End If


        cs.MoveNext
        Loop
 
End Sub

Sub Busca_Cita()
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM CITAS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

            Aleatorio = CLng((1 - totreg) * Rnd + totreg)
        'MsgBox (totreg & "-" & Aleatorio & "-" & numreg)

    numreg = 1
    Do Until cd.EOF = True

        If Aleatorio < numreg Then
            LblNreg = numreg
            LblCita1.Caption = cd.Fields("CITA1")
            PrvCita1 = cd.Fields("CITA1")
            If cd.Fields("CITA2") > "" Then
                LblCita2.Caption = cd.Fields("CITA2")
                PrvCita2 = cd.Fields("CITA2")
            Else
                PrvCita2 = ""
            End If
            If cd.Fields("CITA3") > "" Then
                LblCita3.Caption = cd.Fields("CITA3")
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





Private Sub InicioCaptura_Click()
IntRespuesta = MsgBox("¡¡¡CUIDADO!!!       ESTE PROCESO BORRA LOS REGISTROS DE LA BASE DATOS DE LOS MOVIMIENTOS DEL CUENTAHABIENTE...¡USO FUTURO", 1)
    Exit Sub
If (IntRespuesta = 1) Then
    IntRespuesta = MsgBox("Continúa", 0)
Else
    Exit Sub
End If

IntRespuesta = MsgBox("Data Source=c:\" & Carpeta & "\" & dBanco & ".mdb", 0)

IntRespuesta = MsgBox("¿ESTA SEGURO DE INICIALIZAR REGISTROS PARA CAPTURA...?", 1)
If (IntRespuesta = 1) Then
    IntRespuesta = MsgBox("Continúa", 0)
Else
    Exit Sub
End If
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection
    IntRespuesta = MsgBox("BORRA SICMOV", 0)

    'Borra Registros de SICMOV
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\" & dBanco & ".mdb"
   

   Dim cs As New ADODB.Recordset 'Creamos el Objeto Recordset.SICMOV

   
Set cs = New ADODB.Recordset

    With cs
        .ActiveConnection = cn
        .Source = "SELECT * FROM SICMOV"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    If cs.EOF = True Then
        IntRespuesta = MsgBox("SICMOV NO CONTIENE REGISTROS", 0)
    Else
        cs.MoveFirst
    End If
    
    PrvNumvos = 0
    Do Until cs.EOF = True
        PrvNumovs = PrvNumovs + 1
        cs.Delete
        cs.MoveNext
        
    Loop

End Sub

Private Sub lblFedamac_Click()
    Static lfrmCount As Long
    Dim frmD As FR
    lfrmCount = lfrmCount + 1
    Set frmD = New FR
    frmD.Caption = "FR"
    
    frmD.Show
End Sub

Private Sub MnuAgenda_Click()

frmMiPrimera.Flg = "1"

Static lfrmCount As Long
    Dim frmD As FG
    lfrmCount = lfrmCount + 1
    Set frmD = New FG
    frmD.Caption = "FG"
    
    frmD.Show

End Sub

Private Sub mnuAnual_Click()
    Static lfrmCount As Long
    Dim frmD As FrmCapitalizar
    lfrmCount = lfrmCount + 1
    Set frmD = New FrmCapitalizar
    frmD.Caption = "FrmCapitalizar"
    
    frmD.Show
End Sub

Private Sub mnuArchivoSalir_Click()
    End
End Sub

Private Sub MnuCapDBBMX_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

cBanco = "DBBMX"
LblcBanco = cBanco
Verify_cBanco
End Sub

Private Sub MnuCapDBGMF_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

cBanco = "DBGMF"
LblcBanco = cBanco
Verify_cBanco
End Sub

Private Sub MnuCapDBLLB_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

cBanco = "DBLLB"
LblcBanco = cBanco
Verify_cBanco
End Sub

Private Sub MnuCapDBVCL_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

cBanco = "DBVCL"
LblcBanco = cBanco
Verify_cBanco
End Sub

Private Sub mnucaptura_Click()
    Static lfrmCount As Long
    Dim frmD As FrmCaptura
    lfrmCount = lfrmCount + 1
    Set frmD = New FrmCaptura
    frmD.Caption = "frmCaptura"
    
    frmD.Show
    'Provisional = "P"

   'IntRespuesta = MsgBox("Provisional FrmMENUSYS =" & Provisional, 0)

End Sub

Private Sub Socios()

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   Dim rs As New ADODB.Recordset     'Creamos el objeto Recordset.Sicmov
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim rsFecha As Date
   
   Dim Reg As Integer
   Dim ultreg As Integer
   
   
   'Abrimos la base de datos "sisfed.mdb".
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   rs.Source = "sicmov"        'Especificamos la fuente de datos. En este caso la tabla "sicmov".
   rs.Open "select * from sicmov", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
   cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"
   cl.Open "select * from SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   'rs.Sort by rs.Fields("SOCIO")
   
   'SELECT * FROM SICMOV ORDER BY SOCIO
   'SELECT * FROM "SICMOV" ORDER BY "SOCIO"
    
    Do While Not rs.EOF
       rsimporte = rs.Fields("IMPORTE")
       ultreg = rs.Fields("SOCIO")
       rsFecha = rs.Fields("FECHA")
       rsCtaBco = rs.Fields("CTABCO")
       rs.MoveNext
    Loop
    cl.MoveFirst
        
   Do Until cl.EOF = True
       If cl.Fields("SOCIO") = ultreg Then
          clNombre = cl.Fields("NOMBRE")
          Exit Do
       End If
       cl.MoveNext
    Loop

    'rs.MovePrevious

    'IntRespuesta = MsgBox(ultreg & ".-" & rsimporte & ".-" & clnombre, 1)

      
    Txtultimport.Text = "Socio: " & ultreg & ".-$" & rsimporte & ".-" & clNombre

     'rs.Update (rs.Fields("SOCIO" = "15"))

     'rs.AddNew (rs.Fields("SOCIO" = "15"))
     TxtrsFecha.Text = rsFecha
     TxtrsCtaBco.Text = rsCtaBco




End Sub

Private Sub mnuConSocios_Click()
Static lfrmCount As Long
    Dim frmD As FrmSocios
    lfrmCount = lfrmCount + 1
    Set frmD = New FrmSocios
    frmD.Caption = "FrmSocios"
    
    frmD.Show
End Sub

Private Sub MnuConsultaDBBMF_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

    frmMENUSYS.LblcBanco = "DBGMF"
    ShowMovs
End Sub

Private Sub MnuConsultaDBBMX_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

    frmMENUSYS.LblcBanco = "DBBMX"
    ShowMovs

End Sub

Private Sub MnuConsultaDBGNL_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

    frmMENUSYS.LblcBanco = "DBGNL"
    ShowMovs
End Sub

Private Sub MnuConsultaDBJBS_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

    frmMENUSYS.LblcBanco = "DBJBS"
    ShowMovs
End Sub

Private Sub MnuConsultaDBLLB_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

    frmMENUSYS.LblcBanco = "DBLLB"
    ShowMovs
End Sub

Private Sub MnuConsultaDBVCL_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

    frmMENUSYS.LblcBanco = "DBVCL"
    ShowMovs
End Sub

Private Sub MnuDBGNL_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

cBanco = "DBGNL"
LblcBanco = cBanco
Verify_cBanco
End Sub
Private Sub ShowMovs()
    Static lfrmCount As Long
    Dim frmD As FG
    lfrmCount = lfrmCount + 1
    Set frmD = New FG
    frmD.Caption = "FG"
        
    frmD.Show
End Sub
 
Private Sub Verify_cBanco()
IntRespuesta = MsgBox("CAPTURA BASE DE DATOS " & cBanco, 0)
'Verifica que esta base de datos NO fué capturada previamente
'Con base en el primer registro: compara cBanco VS SISFED (SICMOV)
'Abre Base de Datos SISFED y cBanco
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection SISFED
    
    Dim ch As New ADODB.Recordset 'Creamos el objeto Recordset.SICMOV

    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\" & cBanco & ".mdb"
    ch.Open "SELECT * FROM SICMOV ORDER BY Id", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    If ch.EOF = True Then
      IntRespuesta = MsgBox(cBanco & " ESTA BASE DE DATOS NO CONTIENE REGISTROS", 0)
      Exit Sub
    Else
       ch.MoveFirst
    End If
    'Guarda datos del primer registro para comparación VS SISFED
    cFecha = ch.Fields("FECHA")
    cBco = ch.Fields("CTABCO")
    cImporte = ch.Fields("IMPORTE")
    
    'Abre SICMOV DE SISFED PARA COMAPARACIÓN
    Dim cr As New ADODB.Connection        'Creamos el objeto Connection SISFED
    Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.SICMOV

    cr.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

    cd.Open "SELECT * FROM SICMOV ORDER BY Id", cr  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cd.MoveFirst

    Do Until cd.EOF = True
        If cFecha = cd.Fields("FECHA") And cd.Fields("CTABCO") = cBco And cd.Fields("IMPORTE") = cImporte Then
            IntRespuesta = MsgBox(cFecha & " " & cBco & " " & cImporte & " YA EXISTE EN EL ARCHIVO SICMOV DE SISFED", 0)
            IntRespuesta = MsgBox(ch.Fields("FECHA") & " " & ch.Fields("CTABCO") & " " & ch.Fields("IMPORTE") & " YA EXISTE EN EL ARCHIVO SICMOV DE " & cBanco, 0)
            Exit Sub
        End If
        cd.MoveNext
    Loop
    cd.Close
    
    IntRespuesta = MsgBox("Deseas continuar con el Proceso de Captura de " & cBanco & "...?", 1)
    If (IntRespuesta = 1) Then
        IntRespuesta = MsgBox("Continúa", 0)
    Else
        Exit Sub
    End If
    ch.MoveFirst
 strPath = "C:\" & Carpeta & "\SISFED.mdb"

'Create a new ADO Connection to Northwind
'by using Access and the Jet OLE DB
'provider.

Set cn = New ADODB.Connection

With cn
    .Provider = "Microsoft.Access.OLEDB.10.0"
    .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Data Source").Value = strPath
    .Open
End With

'Create a new ADO Recordset by using a server-side
'keyset cursor and optimistic locking.

Set cs = New ADODB.Recordset

With cs
    .ActiveConnection = cn
    .Source = "SELECT * FROM SICMOV"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CursorLocation = adUseServer
    .Open
End With

Do Until ch.EOF = True
        s_numreg = cs.Fields("Id")
        cGrupo = ch.Fields("GRUPO")
        cSocio = ch.Fields("SOCIO")
        cImporte = ch.Fields("IMPORTE")
        cFecha = ch.Fields("FECHA")
        cCveMov = ch.Fields("CVEMOV")
        cTipo = ch.Fields("TIPO")
        cAPrePac = ch.Fields("APREPAC")
        cDescrip = ch.Fields("DESCRIP")
        cCtaBco = ch.Fields("CTABCO")
        cReferenc = ch.Fields("REFERENC")
        cNumes = ch.Fields("NUMES")
        cTasa = ch.Fields("TASA")
        
    With cs
        .AddNew
        cs.Fields("NUMREG") = cs.Fields("Id")
        s_numreg = cs.Fields("Id")
        cs.Fields("GRUPO") = cGrupo
        cs.Fields("SOCIO") = cSocio
        cs.Fields("IMPORTE") = cImporte
        cs.Fields("FECHA") = cFecha
        cs.Fields("CVEMOV") = cCveMov
        cs.Fields("TIPO") = cTipo
        cs.Fields("APREPAC") = cAPrePac
        cs.Fields("DESCRIP") = cDescrip
        cs.Fields("CTABCO") = cCtaBco
        cs.Fields("REFERENC") = cReferenc
        cs.Fields("NUMES") = cNumes
        cs.Fields("TASA") = cTasa
        cs.Update
    End With
    If cAPrePac = "P" Or cAPrePac = "C" Then
        Actualiza_SALDO
        GrabaDMOVPR
    Else
        Actualiza_SALDO
        GrabaDMOVIN
    End If
    'IntRespuesta = MsgBox(cs.Fields("FECHA") & " " & cs.Fields("CTABCO") & " " & cs.Fields("IMPORTE") & " Datos de la Base de Datos DE SISFED", 0)
    'IntRespuesta = MsgBox(ch.Fields("FECHA") & " " & ch.Fields("CTABCO") & " " & ch.Fields("IMPORTE") & " Datos de la Base de Datos DE " & cBanco, 0)
    'IntRespuesta = MsgBox(cFecha & " " & cCtaBco & " " & cImporte & " Datos de Variables", 0)

    ch.MoveNext
Loop
cs.Close
Static lfrmCount As Long
    Dim frmD As FrmBalance
    lfrmCount = lfrmCount + 1
    Set frmD = New FrmBalance
    frmD.Caption = "frmBalance"
    
    frmD.Show
End Sub
Sub GrabaDMOVIN()
Dim cn As ADODB.Connection
Dim cs As ADODB.Recordset
Dim strPath As String
   
'Update the following path to point to the sample
'Northwind.mdb database on your computer.

strPath = "C:\" & Carpeta & "\SISFED.mdb"

'Create a new ADO Connection to Northwind
'by using Access and the Jet OLE DB
'provider.

Set cn = New ADODB.Connection

With cn
    .Provider = "Microsoft.Access.OLEDB.10.0"
    .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Data Source").Value = strPath
    .Open
End With

'Create a new ADO Recordset by using a server-side
'keyset cursor and optimistic locking.

Set cs = New ADODB.Recordset

With cs
    .ActiveConnection = cn
    .Source = "SELECT * FROM DMOVIN"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CursorLocation = adUseServer
    .Open
End With
Do Until cs.EOF = True
    With cs
        .AddNew
        cs.Fields("NUMREG") = s_numreg
        cs.Fields("GRUPO") = PrvGrupo
        cs.Fields("SOCIO") = cSocio
        cs.Fields("IMPORTE") = cImporte
        cs.Fields("FECHA") = cFecha
        cs.Fields("CVEMOV") = cCveMov
        cs.Fields("TIPO") = cTipo
        cs.Fields("APREPAC") = cAPrePac
        cs.Fields("DESCRIP") = cDescrip
        cs.Fields("CTABCO") = cCtaBco
        cs.Fields("REFERENC") = cReferenc
        cs.Fields("NUMES") = cNumes
        cs.Fields("TASA") = cTasa

        cs.Update
        Exit Do
    End With
    Loop
End Sub
Sub GrabaDMOVPR()
Dim cn As ADODB.Connection
Dim cs As ADODB.Recordset
Dim strPath As String
   
'Update the following path to point to the sample
'Northwind.mdb database on your computer.

strPath = "C:\" & Carpeta & "\SISFED.mdb"

'Create a new ADO Connection to Northwind
'by using Access and the Jet OLE DB
'provider.

Set cn = New ADODB.Connection

With cn
    .Provider = "Microsoft.Access.OLEDB.10.0"
    .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Data Source").Value = strPath
    .Open
End With

'Create a new ADO Recordset by using a server-side
'keyset cursor and optimistic locking.

Set cs = New ADODB.Recordset

With cs
    .ActiveConnection = cn
    .Source = "SELECT * FROM DMOVPR"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CursorLocation = adUseServer
    .Open
End With
Do Until cs.EOF = True
    With cs
        .AddNew
        cs.Fields("NUMREG") = s_numreg
        cs.Fields("GRUPO") = PrvGrupo
        cs.Fields("SOCIO") = cSocio
        cs.Fields("IMPORTE") = cImporte
        cs.Fields("FECHA") = cFecha
        cs.Fields("CVEMOV") = cCveMov
        cs.Fields("TIPO") = cTipo
        cs.Fields("APREPAC") = cAPrePac
        cs.Fields("DESCRIP") = cDescrip
        cs.Fields("CTABCO") = cCtaBco
        cs.Fields("REFERENC") = cReferenc
        cs.Fields("NUMES") = cNumes
        cs.Fields("TASA") = PrvTasa

        cs.Update
        Exit Do
    End With
    Loop
End Sub
Sub Actualiza_SALDO()
    Dim cn As ADODB.Connection
    Dim cl As ADODB.Recordset
    Dim strPath As String
   
    'Update the following path to point to the sample
    'Northwind.mdb database on your computer.

    strPath = "C:\" & Carpeta & "\SISFED.mdb"

    'Create a new ADO Connection to Northwind
    'by using Access and the Jet OLE DB
    'provider.

    Set cn = New ADODB.Connection

    With cn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

    'Create a new ADO Recordset by using a server-side
    'keyset cursor and optimistic locking.

    Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM SOCIOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = cSocio Then
            PrvGrupo = cl.Fields("GRUPO")
            SaldoPres = cl.Fields("SALDOPRES")
            Saldo = cl.Fields("SALDO")
            If cAPrePac = "P" Then
                cl.Fields("SALDOPRES") = SaldoPres - cImporte
                cl.Fields("PAGOS") = cl.Fields("PAGOS") + cImporte
            End If
            If cAPrePac = "C" Then
                cl.Fields("SALDOPRES") = SaldoPres + cImporte
                cl.Fields("PRESTAMOS") = cl.Fields("PRESTAMOS") + cImporte
            End If
            If cAPrePac = "A" Then
                cl.Fields("SALDO") = Saldo + cImporte
                cl.Fields("APORTA") = cl.Fields("APORTA") + cImporte
            End If
            If cAPrePac = "R" Then
                cl.Fields("SALDO") = Saldo - cImporte
                cl.Fields("RETIROS") = cl.Fields("RETIROS") + cImporte
            End If

            cl.Update

            Exit Do
    End If
    cl.MoveNext
    Loop

    'ACTUALIZA SALDO EN CAJA
    cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = "99" Then
            
            If cAPrePac = "P" Then
                cl.Fields("APORTA") = cl.Fields("APORTA") + cImporte
                cl.Fields("SALDO") = cl.Fields("SALDO") + cImporte
            End If
            If cAPrePac = "C" Then
                cl.Fields("RETIROS") = cl.Fields("RETIROS") + cImporte
                cl.Fields("SALDO") = cl.Fields("SALDO") - cImporte
            End If
            If cAPrePac = "A" Then
                cl.Fields("APORTA") = cl.Fields("APORTA") + cImporte
                cl.Fields("SALDO") = cl.Fields("SALDO") + cImporte
            End If
            If cAPrePac = "R" Then
                cl.Fields("RETIROS") = cl.Fields("RETIROS") + cImporte
                cl.Fields("SALDO") = cl.Fields("SALDO") - cImporte
            End If
            'IntRespuesta = MsgBox("SALDO 99=$" & cl.Fields("SALDO"), 0)

            cl.Update
            Exit Do
    End If
    cl.MoveNext
        
    Loop
End Sub

Private Sub MnuDBJBS_Click()
MsgBox ("USO FUTURO...NO APLICA")
Exit Sub

cBanco = "DBJBS"
LblcBanco = cBanco

Verify_cBanco
End Sub

Private Sub mnuEctaGrupo_Click()
    frmMiPrimera.Flg = "0"

    Static lfrmCount As Long
    Dim frmD As MS
    lfrmCount = lfrmCount + 1
    Set frmD = New MS
    frmD.Caption = "MS"
    
    frmD.Show
End Sub



Private Sub mnuEctaSocio_Click()
 
    
      IntRespuesta = MsgBox("MODO DE PRUEBA -NO DISPONIBLE-", 0)

End Sub



Private Sub mnuEctaInv_Click()
  Dim cr As New ADODB.Connection
   
   Dim cd As New ADODB.Recordset

   Dim strPath As String

   cr.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisrpt.mdb"
   
    'Update the following path to point to the sample
    'Northwind.mdb database on your computer.

    strPath = "C:\" & Carpeta & "\SISRPT.mdb"

    'Create a new ADO Connection to Northwind
    'by using Access and the Jet OLE DB
    'provider.

    Set cr = New ADODB.Connection

    With cr
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

    'Create a new ADO Recordset by using a server-side
    'keyset cursor and optimistic locking.

    Set cd = New ADODB.Recordset

    With cd
        .ActiveConnection = cr
        .Source = "SELECT * FROM ECTAINVERSION"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim cp As New ADODB.Recordset


   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   cp.Open "SELECT * FROM DMOVIN ORDER BY GRUPO,SOCIO,FECHA,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cd.MoveFirst
    'Borra Registros del Archivo ECTAINVERSION
    Do Until cd.EOF = True
        x = x + 1
        Me.CurrentX = x
        Me.CurrentY = 2000
        Me.Print ">"
        
        If Not cd.EOF Then
            cd.Delete
            cd.MoveNext
        End If
    Loop
 
 'Graba Resgistros de Movimientos de Inversión DMOVIN
 cp.MoveFirst
 x = 1
 'PrvSocio = cp.Fields("SOCIO")
 Do Until cp.EOF = True
    
   If cp.Fields("IMPORTE") <> 0 Then
    
      With cd
        If cp.Fields("SOCIO") <> PrvSocio Then
            PrvSocio = cp.Fields("SOCIO")
            Busca_Cita

            'IntRespuesta = MsgBox("BUSCA_SOCIO=" & PrvSocio, 0)
            BUSCA_SOCIO
            If Linea <> 0 Then
                'TOTALIZAR COLUMNAS DE APORTACION Y RETIROS
                
                Linea = Linea + 1
                nreg = nreg + 1
                .AddNew
                cd.Fields("Id") = nreg
                cd.Fields("Lin") = 1
                cd.Fields("DESCRIPCION") = "TOTALES"
                cd.Fields("APORTACION") = Format(TotAbonos, "Currency")
                cd.Fields("RETIRO") = Format(TotCargos, "Currency")
                TotAbonos = 0
                TotCargos = 0
                
                Linea = Linea + 1
                nreg = nreg + 1
                .AddNew
                cd.Fields("Id") = nreg
                cd.Fields("Lin") = 1
                cd.Fields("DESCRIPCION") = PrvCita1
                
                Linea = Linea + 1
                nreg = nreg + 1
                .AddNew
                cd.Fields("Id") = nreg
                cd.Fields("Lin") = 1
                cd.Fields("DESCRIPCION") = PrvCita2
                
                Linea = Linea + 1
                nreg = nreg + 1
                .AddNew
                cd.Fields("Id") = nreg
                cd.Fields("Lin") = 1
                cd.Fields("DESCRIPCION") = PrvCita3
                
                Do While Linea < 37
                    Linea = Linea + 1
                    nreg = nreg + 1
                    .AddNew
                    cd.Fields("Id") = nreg
                Loop
            End If
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("Lin") = 1
            cd.Fields("FECHA") = "Socio.-" & cp.Fields("SOCIO")
            cd.Fields("DESCRIPCION") = Left(PubNombre, 30)
            cd.Fields("APORTACION") = "Grupo.-" & cp.Fields("GRUPO")
            'cd.Fields("RETIRO") = "Fecha de"
            cd.Fields("SALDO") = "Fecha de Corte"
            cd.Fields("REFERENCIA") = Format(PrvFecorte, "ddddd")
            cd.Update
             
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "          RESUMEN DE INVERSION"
            cd.Update
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "<-Saldo Inicial           Aportaciones->"
            cd.Fields("SALDO") = "    Promedio->"
            cd.Fields("REFERENCIA") = Format(prvPromedio, "Currency")
            cd.Fields("FECHA") = Format(PrvInvini, "Currency")
            cd.Fields("APORTACION") = Format(PrvAporta, "Currency")
            'cd.Fields("SALDO") = "Fecha Prestamo->"
            'cd.Fields("REFERENCIA") = Format(PrvFecPres, "ddddd")
            cd.Update
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "<-Saldo Actual                 Retiros->"
            cd.Fields("FECHA") = Format(PrvSaldo, "Currency")
            cd.Fields("APORTACION") = Format(PrvRetiros, "Currency")
            If prvIntGanado > 0 Then
                cd.Fields("SALDO") = "Tasa Inversión->"
                cd.Fields("REFERENCIA") = Format(prvIntGanado / prvPromedio, "Percent")
            End If
            cd.Update
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "<-Ints Devengados           Comisiones->"
            cd.Fields("FECHA") = Format(prvIntGanado, "Currency")
            cd.Fields("APORTACION") = Format(prvComision, "Currency")
            'cd.Fields("PAGOS") = Format(PrvPagoMin, "Currency")
            'cd.Fields("SALDO") = "Vencimiento->"
            'cd.Fields("REFERENCIA") = Format(PrvFecVenc, "ddddd")
            
            'LINEA DE PAGO TOTAL
            'Linea = Linea + 1
            'nreg = nreg + 1
            '.AddNew
            'cd.Fields("Id") = nreg
            'cd.Fields("DESCRIPCION") = "                         Pago Total->"
            'cd.Fields("PRESTAMOS") = Format(s_saldopres * (1 + s_tasapres / 100), "Currency")
            'd.Fields("PAGOS") = Format(s_saldopres * (1 + s_tasapres / 100), "Currency")
            
            cd.Update
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "-------------------------------"
            cd.Update

            sdoActual = 0
            Linea = 1
        End If
        If Linea = 37 Then
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("Lin") = 1
            cd.Fields("FECHA") = "Socio.-" & cp.Fields("SOCIO")
            cd.Fields("DESCRIPCION") = Left(PubNombre, 30)
            cd.Fields("APORTACION") = "Grupo.-" & cp.Fields("GRUPO")
            cd.Fields("SALDO") = "Fecha de Corte"
            cd.Fields("REFERENCIA") = Format(PrvFecorte, "ddddd")
            cd.Update
             
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "          RESUMEN DE INVERSION"
            cd.Update
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "<-Saldo Inicial           Aportaciones->"
            cd.Fields("FECHA") = Format(PrvInvini, "Currency")
            cd.Fields("APORTACION") = Format(PrvAporta, "Currency")
            cd.Update
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "<-Saldo Actual                 Retiros->"
            cd.Fields("FECHA") = Format(PrvSaldo, "Currency")
            cd.Fields("APORTACION") = Format(PrvRetiros, "Currency")
            cd.Update
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "<-Ints Devengados           Comisiones->"
            cd.Fields("FECHA") = Format(prvIntGanado, "Currency")
            cd.Fields("APORTACION") = Format(prvComision, "Currency")
            
            cd.Update
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "-------------------------------"
            cd.Update
            Linea = 1
        End If
        
        Linea = Linea + 1
        nreg = nreg + 1
        .AddNew
        cd.Fields("Id") = nreg
        cd.Fields("SOCIO") = cp.Fields("SOCIO")
        cd.Fields("FECHA") = Format(cp.Fields("FECHA"), "ddddd")
        cd.Fields("DESCRIPCION") = cp.Fields("DESCRIP")
        cd.Fields("REFERENCIA") = cp.Fields("REFERENC")
        If cp.Fields("CTABCO") <> "" Then
            cd.Fields("REFERENCIA") = cp.Fields("CTABCO")
        End If
        If cp.Fields("APREPAC") = "A" Then
                  '*Abonos
            cd.Fields("APORTACION") = Format(cp.Fields("IMPORTE"), "Currency")
            TotAbonos = TotAbonos + cp.Fields("IMPORTE")
            sdoActual = sdoActual + cp.Fields("IMPORTE")
        Else
            '      *Cargos
            cd.Fields("RETIRO") = Format(cp.Fields("IMPORTE"), "Currency")
            sdoActual = sdoActual - cp.Fields("IMPORTE")
            TotCargos = TotCargos + cp.Fields("IMPORTE")
        End If
        cd.Fields("SALDO") = Format(sdoActual, "Currency")
        'cd.Fields("REFERENCIA") = cp.Fields("REFERENC")

        cd.Update
      End With
   End If
   cp.MoveNext
    x = x + 5
    Me.CurrentX = x
    Me.CurrentY = 2200
    Me.Print ">>"
 Loop
 Unload Me
        Static lfrmCount As Long
    Dim frmD As frmMENUSYS
    lfrmCount = lfrmCount + 1
    Set frmD = New frmMENUSYS
    frmD.Caption = "frmMENUSYS"
    
    frmD.Show

      'IntRespuesta = MsgBox("MODO DE PRUEBA -NO DISPONIBLE-", 0)

   IntRespuesta = MsgBox("Se generó Estado de Cuenta de Préstamos: ECTAINVERSION en SYSRPT", 0)


End Sub









Private Sub MnuEdoCta_Click()
   Dim cr As New ADODB.Connection
   
   Dim cd As New ADODB.Recordset

   Dim strPath As String

   cr.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisrpt.mdb"
   
    'Update the following path to point to the sample
    'Northwind.mdb database on your computer.

    strPath = "C:\" & Carpeta & "\SISRPT.mdb"

    'Create a new ADO Connection to Northwind
    'by using Access and the Jet OLE DB
    'provider.

    Set cr = New ADODB.Connection

    With cr
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

    'Create a new ADO Recordset by using a server-side
    'keyset cursor and optimistic locking.

    Set cd = New ADODB.Recordset

    With cd
        .ActiveConnection = cr
        .Source = "SELECT * FROM ECTAPRESTAMO"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim cp As New ADODB.Recordset


   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   cp.Open "SELECT * FROM DMOVPR ORDER BY GRUPO,SOCIO,FECHA,APREPAC DESC,TIPO DESC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
   
    
    'Borra Registros del Archivo ECTAPRESTAMO
    Do Until cd.EOF = True
        x = x + 1
        Me.CurrentX = x
        Me.CurrentY = 2000
        Me.Print ">>>"
        cd.MoveFirst
        If Not cd.EOF Then
            cd.Delete
        End If
    Loop
 
 'Graba Resgistros de Movimientos de Préstamos DMOVPR
 cp.MoveFirst
 'PrvSocio = cp.Fields("SOCIO")
 Linea = 0
 x = 1
 Do Until cp.EOF = True
        x = x + 5
        Me.CurrentX = x
        Me.CurrentY = 2200
        Me.Print ">>>"
   If cp.Fields("IMPORTE") <> 0 Then
    
      With cd
        If cp.Fields("SOCIO") <> PrvSocio Then
            PrvSocio = cp.Fields("SOCIO")
            Busca_Cita
            'IntRespuesta = MsgBox("BUSCA_SOCIO=" & SOCIO & "-" & Linea, 0)

            BUSCA_SOCIO
            If Linea <> 0 Then
                'TOTALIZAR COLUMNAS DE PRESTAMOS Y PAGOS
                Linea = Linea + 1
                nreg = nreg + 1
                .AddNew
                cd.Fields("Id") = nreg
                cd.Fields("Lin") = 1
                cd.Fields("DESCRIPCION") = "TOTALES"
                cd.Fields("PAGOS") = Format(TotAbonos, "Currency")
                cd.Fields("PRESTAMOS") = Format(TotCargos, "Currency")
                TotAbonos = 0
                TotCargos = 0
                
                Linea = Linea + 1
                nreg = nreg + 1
                .AddNew
                cd.Fields("Id") = nreg
                cd.Fields("Lin") = 1
                cd.Fields("DESCRIPCION") = PrvCita1
                
                Linea = Linea + 1
                nreg = nreg + 1
                .AddNew
                cd.Fields("Id") = nreg
                cd.Fields("Lin") = 1
                cd.Fields("DESCRIPCION") = PrvCita2
                
                Linea = Linea + 1
                nreg = nreg + 1
                .AddNew
                cd.Fields("Id") = nreg
                cd.Fields("Lin") = 1
                cd.Fields("DESCRIPCION") = PrvCita3
                
                Do While Linea < 37
                    Linea = Linea + 1
                    nreg = nreg + 1
                    .AddNew
                    cd.Fields("Id") = nreg
                Loop
            End If
                
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("Lin") = 1
            cd.Fields("FECHA") = "Socio.-" & cp.Fields("SOCIO")
            cd.Fields("DESCRIPCION") = Left(PubNombre, 30)
            cd.Fields("PRESTAMOS") = "Grupo.-" & cp.Fields("GRUPO")
            cd.Fields("SALDO") = "Fecha de Corte->"
            cd.Fields("REFERENCIA") = Format(PrvFecorte, "ddddd")
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "           RESUMEN DE PRESTAMOS"
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "<-Saldo Inicial              Prestamos->"
            cd.Fields("FECHA") = Format(PrvSaldoInicial, "Currency")
            cd.Fields("PRESTAMOS") = Format(prvPrestamos, "Currency")
            cd.Fields("SALDO") = "Fecha Prestamo->"
            cd.Fields("REFERENCIA") = Format(PrvFecPres, "ddddd")
'Private PrvSaldoInicial, PrvPrestamos As Single
'Private PrvAporta, PrvRetiros, PrvIntPagado, PrvComision, PrvPagos, PrvIntGanado As Single
            

            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "<-Saldo Actual                  Pagos->"
            cd.Fields("FECHA") = Format(s_saldopres, "Currency")
            cd.Fields("PRESTAMOS") = Format(PrvPagos, "Currency")
            'cd.Fields("SALDO") = "Vencimiento->"
            'cd.Fields("REFERENCIA") = Format(PrvFecVenc, "ddddd")
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "<-IntPagado             Pago Minimo->"
            cd.Fields("FECHA") = Format(prvIntPagado, "Currency")
            cd.Fields("PRESTAMOS") = Format(PrvPagoMin, "Currency")
            'cd.Fields("PAGOS") = Format(PrvPagoMin, "Currency")
            cd.Fields("SALDO") = "Vencimiento->"
            cd.Fields("REFERENCIA") = Format(PrvFecVenc, "ddddd")
            
            'LINEA DE PAGO TOTAL
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "                         Pago Total->"
            cd.Fields("PRESTAMOS") = Format(s_saldopres * (1 + s_tasapres / 100), "Currency")
            'd.Fields("PAGOS") = Format(s_saldopres * (1 + s_tasapres / 100), "Currency")
            
            cd.Update
            
            Linea = Linea + 1
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("DESCRIPCION") = "----------------------------------------"
            cd.Update

            sdoActual = 0
            Linea = 1
        End If
        
        Linea = Linea + 1
        nreg = nreg + 1
        .AddNew
        cd.Fields("Id") = nreg
        cd.Fields("SOCIO") = cp.Fields("SOCIO")
        cd.Fields("FECHA") = Format(cp.Fields("FECHA"), "ddddd")
        cd.Fields("DESCRIPCION") = cp.Fields("DESCRIP")
        cd.Fields("REFERENCIA") = cp.Fields("REFERENC")
        If cp.Fields("CTABCO") <> "" Then
            cd.Fields("REFERENCIA") = cp.Fields("CTABCO")
        End If
        If cp.Fields("APREPAC") = "P" Then
                  '*Abonos
            cd.Fields("PAGOS") = Format(cp.Fields("IMPORTE"), "Currency")
            sdoActual = sdoActual - cp.Fields("IMPORTE")
            TotAbonos = TotAbonos + cp.Fields("IMPORTE")
        Else
            '      *Cargos
            cd.Fields("PRESTAMOS") = Format(cp.Fields("IMPORTE"), "Currency")
            cd.Fields("REFERENCIA") = Format(cp.Fields("TASA") / 100, "Percent")
            sdoActual = sdoActual + cp.Fields("IMPORTE")
            TotCargos = TotCargos + cp.Fields("IMPORTE")
        End If
        cd.Fields("SALDO") = Format(sdoActual, "Currency")
     
        cd.Update
      End With
   End If
   cp.MoveNext

 Loop
 Unload Me
        Static lfrmCount As Long
    Dim frmD As frmMENUSYS
    lfrmCount = lfrmCount + 1
    Set frmD = New frmMENUSYS
    frmD.Caption = "frmMENUSYS"
    
    frmD.Show
      'IntRespuesta = MsgBox("MODO DE PRUEBA -NO DISPONIBLE-", 0)

   IntRespuesta = MsgBox("Se generó Estado de Cuenta de Préstamos: ECTAPRESTAMO en SYSRPT", 0)

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
            'IntRespuesta = MsgBox("BUSCA_SOCIO=" & PrvSocio, 0)
            If cl.Fields("FECVENC") <> "" Then
                PrvMeses = (cl.Fields("FECVENC") - cl.Fields("FECPRES")) / 30.4
            End If
            PubNombre = cl.Fields("NOMBRE")
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
            If cl.Fields("PROM_INV") <> "" Then
                prvPromedio = cl.Fields("PROM_INV")
            End If

            
            Exit Do
        End If
        cl.MoveNext
        Loop
End Sub
    
Private Sub BUSCA_SOCIO_DBF()
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS1", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
    cl.MoveFirst
    'IntRespuesta = MsgBox("BUSCA_SOCIO=" & PrvSocio, 0)

    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = PrvSocio Then
            'IntRespuesta = MsgBox("BUSCA_SOCIO=" & PrvSocio, 0)
            Saldo_DBF = cl.Fields("SALDO")
            IntGanado_DBF = cl.Fields("INTGANADO")
            Comision_DBF = cl.Fields("COMISION")
            SaldoPres_DBF = cl.Fields("SALDOPRES")
            Exit Do
        End If
        cl.MoveNext
        Loop
End Sub
Private Sub cmd_exportar_click()

                   'Establecemos la ruta de nuestro archivo
                   Ruta = Carpeta & "\orden.doc"

                  'Seteamos el archivo al objeto documento
                  Set Documento = MSWord.Documents.Open(Ruta)

                  'opcionalmente podemos guardar el archivo
                  'en mi caso lo guardo con una extensión diferente (cab|tmp|pot|etc)
                  MSWord.Selection.Document.SaveAs (Carpeta & "\printme.cab")

                  'Establecemos la fuentre que utilizaremos
                  MSWord.Selection.Font.Name = "Arial"

                  'Configuramos la alineacion de nuestro parrafo
                  MSWord.Selection.Paragraphs.Alignment = wdAlignParagraphCenter

                  'Activamos la fuente en Negrita
                  MSWord.Selection.Font.Bold = True

                  'Y el tamaño a 16 puntos
                  MSWord.Selection.Font.Size = 16

                  'con esta opcion podemos comenzar a escribir dentro de nuestro docuemnto
                  MSWord.Selection.TypeText "Aqui podemos escribir el texto en el documento" & vbCrLf

                  'Declaramos una tabla de 1 fila por 3 columnas
                  MSWord.Selection.Tables.Add MSWord.Selection.Range, 1, 3

                  'Seleccionamos la celda 1,2
                  MSWord.Selection.Tables(1).Cell(1, 2).Select

                  'establecemos el ancho de la celda
                 MSWord.Selection.Tables(1).Cell(1, 2).Width = 70

                  'configuramos los bordes
                  MSWord.Selection.Tables(1).Cell(1, 2).Borders(wdBorderTop).Visible = True
                  MSWord.Selection.Tables(1).Cell(1, 2).Borders(wdBorderLeft).Visible = True
                  MSWord.Selection.Tables(1).Cell(1, 2).Borders(wdBorderBottom).Visible = True
                  MSWord.Selection.Tables(1).Cell(1, 2).Borders(wdBorderRight).Visible = True

                  'Y la alineación del texto dentro de la celda
                  MSWord.Selection.Paragraphs.Alignment = wdAlignParagraphLeft

                  'Seguido escribimos texto en dicha celda
                  MSWord.Selection.TypeText "Nombre"

                  'seleccionamos la celda 1,3
                  MSWord.Selection.Tables(1).Cell(1, 3).Select

                  'Establcemos el color de fondo de la celda (Trama)
                  MSWord.Selection.Cells.Shading.BackgroundPatternColor = wdColorGray20

                  'Escribimos en dicha celda
                  MSWord.Selection.TypeText "nombre2"

                  'esta opcion nos permite salir de la edición de la tabla, o bajar una fila
                  MSWord.Selection.MoveDown

                  'por ultimo mostramos el documento de word
                  MSWord.Visible = True

                  'vaciamos los objetos de la  memoria
                  Set Documento = Nothing
                  Set MSWord = Nothing

End Sub

Private Sub MnuInicioEjercicio_Click()
IntRespuesta = MsgBox("¡¡¡CUIDADO!!!       ESTE PROCESO BORRA LOS REGISTROS DEL EJERCICIO ANTERIOR E INICIALIZA LOS DATOS PARA EL SIGUIENTE EJERCICIO", 1)
If (IntRespuesta = 1) Then
    IntRespuesta = MsgBox("Continúa", 0)
Else
    Exit Sub
End If
IntRespuesta = MsgBox("¿ESTA SEGURO DE INICIALIZAR REGISTROS PARA NUEVO EJERCICIO?", 1)
If (IntRespuesta = 1) Then
    IntRespuesta = MsgBox("Continúa", 0)
Else
    Exit Sub
End If
f_final = "31/10/2011"

    'Borra DATOS de SOCIOS
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el Objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM SOCIOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    'Inicializa DATOS del Archivo SOCIOS
    cl.MoveFirst

    Do Until cl.EOF = True
        
        cl.Fields("INV_INI") = cl.Fields("SALDO") + cl.Fields("INTGANADO") + cl.Fields("COMISION")
        cl.Fields("SALDO") = cl.Fields("INV_INI")
        If cl.Fields("SOCIO") = "988" Then
            s_intganado = cl.Fields("INTGANADO")
            cl.Fields("INV_INI") = 0
            cl.Fields("SALDO") = 0
        End If
        If cl.Fields("SOCIO") = "25" Then
            s_saldopres = cl.Fields("SALDOPRES")
            cl.Fields("SALDOPRES") = 0
        End If
        
        cl.Fields("PRES_INI") = cl.Fields("SALDOPRES")
        cl.Fields("APORTA") = 0
        cl.Fields("RETIROS") = 0
        cl.Fields("INTPAGADO") = 0
        cl.Fields("COMISION") = 0
        cl.Fields("PRESTAMOS") = 0
        cl.Fields("PAGOS") = 0
        cl.Fields("INTGANADO") = 0
        cl.Update
        cl.MoveNext
    Loop
    cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = "50" Then
            cl.Fields("SALDO") = cl.Fields("SALDO") + s_intganado - s_saldopres
            cl.Fields("INV_INI") = cl.Fields("SALDO")
            cl.Update
            Exit Do
        End If
        cl.MoveNext
    Loop
'BORRA REGISTROS DE SICMOV
IntRespuesta = MsgBox("BORRA SICMOV", 0)

    'Borra Registros de SICMOV

   Dim cs As New ADODB.Recordset 'Creamos el Objeto Recordset.SICMOV

   
Set cs = New ADODB.Recordset

    With cs
        .ActiveConnection = cn
        .Source = "SELECT * FROM SICMOV"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    cs.MoveFirst
    PrvNumvos = 0
    Do Until cs.EOF = True
        PrvNumovs = PrvNumovs + 1
        cs.Delete
        cs.MoveNext
        
    Loop

IntRespuesta = MsgBox("BORRA REGISTROS DMOVIN", 0)

    'Borra Registros de DMOVIN

   Dim cv As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVIN

   
Set cv = New ADODB.Recordset

    With cv
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVIN"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

Set cv = New ADODB.Recordset

    With cv
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVIN"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    'Borra Registros del Archivo DMOVIN
    cv.MoveFirst

    Do Until cv.EOF = True
        PrvNumovs = PrvNumovs + 1
                cv.Delete
                cv.MoveNext
    Loop
IntRespuesta = MsgBox("BORRA REGISTROS DMOVPR", 0)

    'Borra Registros de DMOVPR

   Dim cp As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVPR

  
    Set cp = New ADODB.Recordset

    With cp
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVPR"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

Set cp = New ADODB.Recordset

    With cp
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVPR"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    'Borra Registros del Archivo DMOVPR
    cp.MoveFirst

    Do Until cp.EOF = True
        PrvNumovs = PrvNumovs + 1
        cp.Delete
        cp.MoveNext
    Loop
    IntRespuesta = MsgBox("INICIALIZA DMOVPR, DMOVIN Y SICMOV CON SALDO INICIAL DE CADA SOCIO", 0)
   cl.Close
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
    
    cl.MoveFirst
    IdNum = 0
    Do Until cl.EOF = True
    With cp
        .AddNew
        If IdNum = 0 Then
            cp.Fields("Id") = 1
        End If
        cp.Fields("NUMREG") = cp.Fields("Id")
        cp.Fields("DESCRIP") = "SALDO ANTERIOR"
        cp.Fields("GRUPO") = cl.Fields("GRUPO")
        cp.Fields("SOCIO") = cl.Fields("SOCIO")
        cp.Fields("IMPORTE") = cl.Fields("PRES_INI")
        cp.Fields("FECHA") = f_final + 1
        cp.Fields("CVEMOV") = "60"
        cp.Fields("TIPO") = "C"
        cp.Fields("APREPAC") = "C"
        cp.Fields("NUMES") = 1
        cp.Update
    End With
    With cv
        .AddNew
        If IdNum = 0 Then
            cv.Fields("Id") = 1
        End If
        cv.Fields("NUMREG") = cv.Fields("Id")
        cv.Fields("DESCRIP") = "SALDO ANTERIOR"
        cv.Fields("GRUPO") = cl.Fields("GRUPO")
        cv.Fields("SOCIO") = cl.Fields("SOCIO")
        cv.Fields("IMPORTE") = cl.Fields("INV_INI")
        cv.Fields("FECHA") = f_final + 1
        cv.Fields("CVEMOV") = "00"
        cv.Fields("TIPO") = "A"
        cv.Fields("APREPAC") = "A"
        cv.Fields("NUMES") = 1
        cv.Update
    End With
    With cs
        .AddNew
        If IdNum = 0 Then
            cs.Fields("Id") = 1
        End If
        cs.Fields("NUMREG") = cs.Fields("Id")
        cs.Fields("DESCRIP") = "SALDO ANTERIOR"
        cs.Fields("GRUPO") = cl.Fields("GRUPO")
        cs.Fields("SOCIO") = cl.Fields("SOCIO")
        cs.Fields("IMPORTE") = cl.Fields("PRES_INI")
        cs.Fields("FECHA") = f_final + 1
        cs.Fields("CVEMOV") = "60"
        cs.Fields("TIPO") = "C"
        cs.Fields("APREPAC") = "C"
        cs.Fields("NUMES") = 1
        cs.Update
    End With
    With cs
        .AddNew
        If IdNum = 0 Then
            cs.Fields("Id") = 2
            IdNum = 1
        End If
        cs.Fields("NUMREG") = cs.Fields("Id")
        cs.Fields("DESCRIP") = "SALDO ANTERIOR"
        cs.Fields("GRUPO") = cl.Fields("GRUPO")
        cs.Fields("SOCIO") = cl.Fields("SOCIO")
        cs.Fields("IMPORTE") = cl.Fields("INV_INI")
        cs.Fields("FECHA") = f_final + 1
        cs.Fields("CVEMOV") = "00"
        cs.Fields("TIPO") = "A"
        cs.Fields("APREPAC") = "A"
        cs.Fields("NUMES") = 1
        cs.Update
    End With
        cl.MoveNext
    Loop
    
    IntRespuesta = MsgBox("TERMINA INICIALIZACION DE EJERCICIO", 0)

End Sub


Private Sub mnuListaNombres_Click()
    Static lfrmCount As Long
    Dim frmD As FrmBalance
    lfrmCount = lfrmCount + 1
    Set frmD = New FrmBalance
    frmD.Caption = "frmBalance"
    
    frmD.Show

End Sub

Private Sub MnuListaPrestamos_Click()
Static lfrmCount As Long
    Dim frmD As FGP
    lfrmCount = lfrmCount + 1
    Set frmD = New FGP
    frmD.Caption = "FGP"
    
    frmD.Show
End Sub

Private Sub mnuMensajeInicial_Click()
    Dim cn As ADODB.Connection
    Dim cl As ADODB.Recordset
    Dim strPath As String
   
    'Update the following path to point to the sample
    'Northwind.mdb database on your computer.

    strPath = "C:\" & Carpeta & "\SISFED.mdb"

    'Create a new ADO Connection to Northwind
    'by using Access and the Jet OLE DB
    'provider.

    Set cn = New ADODB.Connection

    With cn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

    'Create a new ADO Recordset by using a server-side
    'keyset cursor and optimistic locking.

    Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM ECTAFAU"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    cl.MoveLast
   With cl
        .AddNew
        cl.Fields("FECHA") = "Socio.-12"
        cl.Fields("DESCRIPCION") = "CLAUDIA HERMOSILLO BAEZA"
        'cl.Fields("APORTACION") = Format(100, "###,###,##0.00")
        cl.Fields("RETIRO") = "Fecha de"
        'cl.Fields("BANCO") = "MBO"
        cl.Fields("SALDO") = "Corte---->"
        cl.Fields("REFERENCIA") = "16/06/2010"
        cl.Update
   End With
   IntRespuesta = MsgBox("SE AGREGÓ UN REGISTRO", 0)

    Static lfrmCount As Long
    Dim frmD As frmMENUSYS
    lfrmCount = lfrmCount + 1
    Set frmD = New frmMENUSYS
    frmD.Caption = "frmMENUSYS"
    
    frmD.Show

End Sub

Private Sub MnuListaPrestamos1_Click()
frmMiPrimera.Flg = "1"
MnuListaPrestamos_Click
End Sub

Private Sub MnuReembolso_Click()
'IntRespuesta = MsgBox("CAPTURA REEMBOLSO AUTOMATICO", 0)
'Verifica que esta base de datos NO fué capturada previamente
'Con base en el primer registro: compara AGENCDA(Asistencia) VS SISFED(SICMOV)
'Abre Base de Datos SISFED y cBanco
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection SISFED
    
    Dim ch As New ADODB.Recordset 'Creamos el objeto Recordset.SICMOV
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & "SYSVOTA" & "\" & "Agenda" & ".mdb"
    ch.Open "SELECT * FROM ASISTENCIA ORDER BY Id", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    If ch.EOF = True Then
      IntRespuesta = MsgBox(cBanco & " ESTA BASE DE DATOS NO CONTIENE REGISTROS", 0)
      Exit Sub
    Else
       ch.MoveFirst
    End If
    'Guarda datos del primer registro para comparación VS SISFED
    'cFecha = ch.Fields("FECORTE")
    cReembolso = ch.Fields("REEMBOLSO")
    PrvSocio = ch.Fields("SOCIO")
    
    'Abre SICMOV DE SISFED PARA COMAPARACIÓN
    Dim cr As New ADODB.Connection        'Creamos el objeto Connection SISFED
    Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.SICMOV

    cr.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

    cd.Open "SELECT * FROM SICMOV ORDER BY Id", cr  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cd.MoveFirst

    Do Until cd.EOF = True
        If cFecha = cd.Fields("FECHA") And cd.Fields("SOCIO") = PrvSocio And cd.Fields("IMPORTE") = cReembolso Then
            IntRespuesta = MsgBox(cFecha & " " & cBco & " " & cImporte & " YA EXISTE EN EL ARCHIVO SICMOV DE SISFED", 0)
            IntRespuesta = MsgBox(ch.Fields("FECHA") & " " & ch.Fields("CTABCO") & " " & ch.Fields("IMPORTE") & " YA EXISTE EN EL ARCHIVO SICMOV DE " & cBanco, 0)
            Exit Sub
        End If
        cd.MoveNext
    Loop
    cd.Close
    
    IntRespuesta = MsgBox("Deseas continuar con el Proceso de Captura de REEMBOLSO AUTOMATICO" & "...?", 1)
    If (IntRespuesta = 1) Then
        IntRespuesta = MsgBox("Continúa", 0)
    Else
        Exit Sub
    End If
    ch.MoveFirst
 strPath = "C:\" & Carpeta & "\SISFED.mdb"

'Create a new ADO Connection to Northwind
'by using Access and the Jet OLE DB
'provider.

Set cn = New ADODB.Connection

With cn
    .Provider = "Microsoft.Access.OLEDB.10.0"
    .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Data Source").Value = strPath
    .Open
End With

'Create a new ADO Recordset by using a server-side
'keyset cursor and optimistic locking.

Set cs = New ADODB.Recordset

With cs
    .ActiveConnection = cn
    .Source = "SELECT * FROM SICMOV"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CursorLocation = adUseServer
    .Open
End With

Do Until ch.EOF = True
If ch.Fields("REEMBOLSO") > 0 Then
        s_numreg = cs.Fields("Id")
        cGrupo = ch.Fields("GRUPO")
        cSocio = ch.Fields("SOCIO")
        cImporte = ch.Fields("REEMBOLSO")
        cFecha = Date
        'ch.Fields ("FECORTE")
        cCveMov = "15"
        cTipo = "T"
        cAPrePac = "A"
        cDescrip = ch.Fields("COMIDA")
        cReferenc = "JUNTA GRAL"
        If Month(cFecha) > 10 Then
           cNumes = Month(cFecha) - 10
        Else
           cNumes = Month(cFecha) + 2
        End If
        cTasa = 0
        TotReembolso = TotReembolso + cImporte

    With cs
        .AddNew
        cs.Fields("NUMREG") = cs.Fields("Id")
        s_numreg = cs.Fields("Id")
        cs.Fields("GRUPO") = cGrupo
        cs.Fields("SOCIO") = cSocio
        cs.Fields("IMPORTE") = cImporte
        cs.Fields("FECHA") = cFecha
        cs.Fields("CVEMOV") = cCveMov
        cs.Fields("TIPO") = cTipo
        cs.Fields("APREPAC") = cAPrePac
        cs.Fields("DESCRIP") = cDescrip
        cs.Fields("REFERENC") = cReferenc
        cs.Fields("NUMES") = cNumes
        cs.Fields("TASA") = cTasa
        cs.Update
    End With
    If cAPrePac = "P" Or cAPrePac = "C" Then
        Actualiza_SALDO
        GrabaDMOVPR
    Else
        Actualiza_SALDO
        GrabaDMOVIN
    End If
End If
ch.MoveNext

Loop
With cs
        cGrupo = "99"
        cSocio = "50"
        cImporte = TotReembolso
        cCveMov = "33"
        cTipo = "T"
        cAPrePac = "R"
        cDescrip = "REEMBOLSO COMIDA"
        cReferenc = "JUNTA GRAL"
        .AddNew
         cs.Fields("NUMREG") = cs.Fields("Id")
        s_numreg = cs.Fields("Id")
        cs.Fields("GRUPO") = cGrupo
        cs.Fields("SOCIO") = cSocio
        cs.Fields("IMPORTE") = cImporte
        cs.Fields("FECHA") = cFecha
        cs.Fields("CVEMOV") = cCveMov
        cs.Fields("TIPO") = cTipo
        cs.Fields("APREPAC") = cAPrePac
        cs.Fields("DESCRIP") = cDescrip
        cs.Fields("REFERENC") = cReferenc
        cs.Fields("NUMES") = cNumes
        cs.Fields("TASA") = cTasa
        
        cs.Update
    End With
    If cAPrePac = "P" Or cAPrePac = "C" Then
        Actualiza_SALDO
        GrabaDMOVPR
    Else
        Actualiza_SALDO
        GrabaDMOVIN
    End If
cs.Close
Static lfrmCount As Long
    Dim frmD As FrmBalance
    lfrmCount = lfrmCount + 1
    Set frmD = New FrmBalance
    frmD.Caption = "frmBalance"
    
    frmD.Show
End Sub


Private Sub mnurelmov_Click()
'IntRespuesta = MsgBox(pbSocio, 0)
frmMiPrimera.Flg = "0"
Static lfrmCount As Long
    Dim frmD As FG
    lfrmCount = lfrmCount + 1
    Set frmD = New FG
    frmD.Caption = "FG"
    
    frmD.Show
End Sub

Private Sub TxtrsSocio_Change()
Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"
   cl.Open "select * from SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    
   ' cl.MoveFirst
    Importe = Txtrsimporte
    NumSocio = Right(Importe, 4)
    NumSocio = NumSocio * 100
    'TxtrsSocio = NumSocio
    'IntRespuesta = MsgBox(NumSocio, 0)

    Do Until cl.EOF = True
       If cl.Fields("SOCIO") = TxtrsSocio Then
          clNombre = cl.Fields("NOMBRE")
          TxtclNombre = clNombre
          Exit Do
       Else
          TxtclNombre = "No existe nombre de este Socio"

       End If
       cl.MoveNext
    Loop

    'If TxtrsSocio <> " " Then
     '  GoTo BuscaNombre
    'End If
End Sub
Private Sub BuscaNombre()

    cl.MoveFirst
    Do Until cl.EOF = True
       If cl.Fields("SOCIO") = ultreg Then
          clNombre = cl.Fields("NOMBRE")
          Exit Do
       End If
       cl.MoveNext
    Loop
End Sub

Private Sub TxtSN_Change()
    If TxtSN = "N" Then
       IntRespuesta = MsgBox("El movimientos no se grabó", 0)
    End If
    If TxtSN = "S" Then
       IntRespuesta = MsgBox("El movimientos será Grabado cuando se pueda", 1)
    End If
End Sub

Private Sub MnuRelSocios_Click()
   Dim cr As New ADODB.Connection
   
   Dim cd As New ADODB.Recordset

   Dim strPath As String

   cr.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisrpt.mdb"
   
    'Update the following path to point to the sample
    'Northwind.mdb database on your computer.

    strPath = "C:\" & Carpeta & "\SISRPT.mdb"

    'Create a new ADO Connection to Northwind
    'by using Access and the Jet OLE DB
    'provider.

    Set cr = New ADODB.Connection

    With cr
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

    'Create a new ADO Recordset by using a server-side
    'keyset cursor and optimistic locking.

    Set cd = New ADODB.Recordset

    With cd
        .ActiveConnection = cr
        .Source = "SELECT * FROM RSOC"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes


   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   cl.Open "SELECT * FROM SOCIOS ORDER BY TIPO,SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   
   cl.MoveFirst
    Do Until cl.EOF = True
        TotSaldo = TotSaldo + cl.Fields("SALDO")
        cl.MoveNext
        Loop
    'Borra Registros de RSOC
    Do Until cd.EOF = True
        cd.MoveFirst
        If Not cd.EOF Then
            cd.Delete
        End If
    Loop

 'Graba Resgistros de Movimientos de Préstamos RSOC
 cl.MoveFirst
 PrvFecorte = cl.Fields("FECORTE")

 nreg = 0
 TotPorciento = 0
 s_tipo = 4
 Dim tmayor As Single
    Do Until cl.EOF = True
      With cd

    'IntRespuesta = MsgBox("SOCIO=" & cl.Fields("SOCIO"), 0)
        If cl.Fields("TIPO") > s_tipo Then
            s_tipo = cl.Fields("TIPO")
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("NOMBRE") = "TOTALES"
            cd.Fields("INVERSION") = Format(tinversion, "###,###.00")
            cd.Fields("INTGANADO") = Format(tintganado, "###,###.00")
            cd.Fields("COMISION") = Format(tcomision, "###,###.00")
            cd.Fields("INTPAGADO") = Format(tintpagado, "###,###.00")
            cd.Fields("PRESTAMOS") = Format(tprestamos, "###,###.00")
            cd.Fields("PROMEDIO") = Format(tpromedio, "###,###.00")
            cd.Update
            
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Update
            
            ginversion = ginversion + tinversion
            gintganado = gintganado + tintganado
            gcomision = gcomision + tcomision
            gintpagado = gintpagado + tintpagado
            gprestamos = gprestamos + tprestamos
            gpromedio = gpromedio + tpromedio
            
            tinversion = 0
            tintganado = 0
            tcomision = 0
            tintpagado = 0
            tprestamos = 0
            tpromedio = 0
        End If
        If nreg = 0 Then
            nreg = nreg + 1
           .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("NOMBRE") = "FECHA DE CORTE:  " & PrvFecorte
        End If
        nreg = nreg + 1
        .AddNew
        cd.Fields("Id") = nreg
        cd.Fields("TIPO") = cl.Fields("TIPO")
        cd.Fields("SOCIO") = cl.Fields("SOCIO")
        cd.Fields("NOMBRE") = Left(cl.Fields("NOMBRE"), 30)
        cd.Fields("INVERSION") = Format(cl.Fields("SALDO"), "###,###.00")
        cd.Fields("INTGANADO") = Format(cl.Fields("INTGANADO"), "###,###.00")
        cd.Fields("COMISION") = Format(cl.Fields("COMISION"), "###,###.00")
        cd.Fields("INTPAGADO") = Format(cl.Fields("INTPAGADO"), "###,###.00")
        cd.Fields("PRESTAMOS") = Format(cl.Fields("SALDOPRES"), "###,###.00")
        If cl.Fields("PROM_INV") <> "" Then
            cd.Fields("PROMEDIO") = Format(cl.Fields("PROM_INV"), "###,###.00")
        End If
        tinversion = tinversion + cl.Fields("SALDO")
        tintganado = tintganado + cl.Fields("INTGANADO")
        tcomision = tcomision + cl.Fields("COMISION")
        tintpagado = tintpagado + cl.Fields("INTPAGADO")
        tprestamos = tprestamos + cl.Fields("SALDOPRES")
        If cl.Fields("PROM_INV") <> "" Then
            tpromedio = tpromedio + cl.Fields("PROM_INV")
        End If
        cd.Update

        
      End With
            
 
        cl.MoveNext
    Loop
    
    With cd
        nreg = nreg + 1
        .AddNew
        cd.Fields("Id") = nreg
        'cd.Fields("TIPO") = s_tipo
        'cd.Fields("SOCIO") = s_socio
        cd.Fields("NOMBRE") = "TOTALES "
        cd.Fields("INVERSION") = Format(tinversion, "#,###,###.00")
        cd.Fields("INTGANADO") = Format(tintganado, "###,###.00")
        cd.Fields("COMISION") = Format(tcomision, "###,###.00")
        cd.Fields("INTPAGADO") = Format(tintpagado, "###,###.00")
        cd.Fields("PRESTAMOS") = Format(tprestamos, "#,###,###.00")
        cd.Fields("PROMEDIO") = Format(tpromedios, "#,###,###.00")
        cd.Update
    End With
            ginversion = ginversion + tinversion
            gintganado = gintganado + tintganado
            gcomision = gcomision + tcomision
            gintpagado = gintpagado + tintpagado
            gprestamos = gprestamos + tprestamos
            gpromedio = gpromedio + tpromedio
    With cd
        nreg = nreg + 1
        .AddNew
        cd.Fields("Id") = nreg
        'cd.Fields("TIPO") = s_tipo
        'cd.Fields("SOCIO") = s_socio
        cd.Fields("NOMBRE") = "TOTALES GENERALES"
        cd.Fields("INVERSION") = Format(ginversion, "#,###,###.00")
        cd.Fields("INTGANADO") = Format(gintganado, "###,###.00")
        cd.Fields("COMISION") = Format(gcomision, "###,###.00")
        cd.Fields("INTPAGADO") = Format(gintpagado, "###,###.00")
        cd.Fields("PRESTAMOS") = Format(gprestamos, "#,###,###.00")
        cd.Fields("PROMEDIO") = Format(gpromedio, "#,###,###.00")
        cd.Update
    End With
   IntRespuesta = MsgBox("Se generó REPORTE DE SOCIOS PARA EXCEL en DB RSOC en SYSRPT", 0)

  

End Sub

Private Sub MnuRMAY_Click()
      Dim cr As New ADODB.Connection
   
   Dim cd As New ADODB.Recordset

   Dim strPath As String

   cr.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisrpt.mdb"
   
    'Update the following path to point to the sample
    'Northwind.mdb database on your computer.

    strPath = "C:\" & Carpeta & "\SISRPT.mdb"

    'Create a new ADO Connection to Northwind
    'by using Access and the Jet OLE DB
    'provider.

    Set cr = New ADODB.Connection

    With cr
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

    'Create a new ADO Recordset by using a server-side
    'keyset cursor and optimistic locking.

    Set cd = New ADODB.Recordset

    With cd
        .ActiveConnection = cr
        .Source = "SELECT * FROM RMAY"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes


   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   cl.Open "SELECT * FROM SOCIOS ORDER BY SALDO DESC", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   
   cl.MoveFirst
   PrvFecorte = cl.Fields("FECORTE")

    Do Until cl.EOF = True
        TotSaldo = TotSaldo + cl.Fields("SALDO")
        cl.MoveNext
        Loop
    'Borra Registros de RMAY
    Do Until cd.EOF = True
        cd.MoveFirst
        If Not cd.EOF Then
            cd.Delete
        End If
    Loop

 'Graba Resgistros de Movimientos de Préstamos RMAY
 cl.MoveFirst
 nreg = 0
 TotPorciento = 0
 Dim tmayor As Single
    Do Until cl.EOF = True
    'IntRespuesta = MsgBox("SOCIO=" & cl.Fields("SOCIO"), 0)

     If cl.Fields("SALDO") > 0 Then
     If cl.Fields("TIPO") <> "9" Then
      With cd
      
        Porciento = cl.Fields("SALDO") / TotSaldo
        tmayor = tmayor + (Porciento * 100)
        TotPorciento = Porciento + TotPorciento
        If nreg = 0 Then
            nreg = nreg + 1
           .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("NOMBRE") = "FECHA DE CORTE:  " & _
                Format(PrvFecorte, "dd-mmm-yyyy")

        End If
        nreg = nreg + 1
        .AddNew
        cd.Fields("Id") = nreg
        cd.Fields("Por%") = Format(Porciento, "Percent")
        cd.Fields("SOCIO") = cl.Fields("SOCIO")
        cd.Fields("NOMBRE") = Left(cl.Fields("NOMBRE"), 30)
        cd.Fields("INVERSION") = Format(cl.Fields("SALDO"), "###,###.00")
        cd.Fields("INTGANADO") = Format(cl.Fields("INTGANADO"), "###,###.00")
        cd.Fields("COMISION") = Format(cl.Fields("COMISION"), "###,###.00")
        cd.Fields("INTPAGADO") = Format(cl.Fields("INTPAGADO"), "###,###.00")
        cd.Fields("PRESTAMOS") = Format(cl.Fields("SALDOPRES"), "###,###.00")
        cd.Fields("PROMEDIO") = Format(cl.Fields("PROM_INV"), "###,###.00")
        cd.Fields("FECORTE") = cl.Fields("FECORTE")
        cd.Update

        If tmayor > 50 Then
            tmayor = 0
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("Por%") = Format(TotPorciento, "Percent")
            cd.Fields("NOMBRE") = "SOCIOS MAYORITARIOS"
            cd.Update
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Update

        End If
      End With
            tinversion = 0
            tintganado = 0
            tcomision = 0
            tintpagado = 0
            tprestamos = 0
            tpromedio = 0
            sgrupo = cl.Fields("GRUPO")
     End If
     End If
        Socio = cl.Fields("SOCIO")
        tinversion = tinversion + cl.Fields("SALDO")
        tintganado = tintganado + cl.Fields("INTGANADO")
        tcomision = tcomision + cl.Fields("COMISION")
        tintpagado = tintpagado + cl.Fields("INTPAGADO")
        tprestamos = tprestamos + cl.Fields("SALDOPRES")
        If cl.Fields("PROM_INV") <> "" Then
            tpromedio = tpromedio + cl.Fields("PROM_INV")
        End If

        cl.MoveNext
    Loop
    
   IntRespuesta = MsgBox("Se generó REPORTE DE SOCIOS MAYORITARIOS PARA EXCEL en DB RMAY en SYSRPT", 0)


End Sub

Private Sub MnuRpre_Click()
    'Graba Lista de Préstamos efectuados en el Ejercico en una Tabla de GLEX GRID
   Dim cr As New ADODB.Connection
   
   Dim cp As New ADODB.Recordset

   Dim strPath As String

   cr.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisrpt.mdb"
   
    'Update the following path to point to the sample
    'Northwind.mdb database on your computer.

    strPath = "C:\" & Carpeta & "\SISRPT.mdb"

    'Create a new ADO Connection to Northwind
    'by using Access and the Jet OLE DB
    'provider.

    Set cr = New ADODB.Connection

    With cr
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

    'Create a new ADO Recordset by using a server-side
    'keyset cursor and optimistic locking.

    Set cp = New ADODB.Recordset

    With cp
        .ActiveConnection = cr
        .Source = "SELECT * FROM RPRE"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
  
    If Not cp.EOF Then
    
    cp.MoveFirst
    'Borra Registros de RPRE
    Do Until cp.EOF = True
        cp.MoveFirst
        If Not cp.EOF Then
            cp.Delete
        End If
    Loop
    End If
    End With

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   'Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.SICMOV

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   cd.Open "SELECT * FROM SICMOV ORDER BY SOCIO,NUMES", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   

 'Graba Registros de Movimientos de Préstamos RPRE
 cd.MoveFirst
 nreg = 0
 
   'SELECT SICMOV
      s_cvemov = cd.Fields("CVEMOV")
      s_tipo = cd.Fields("TIPO")
      n_socio = cd.Fields("SOCIO")
    PrvSocio = n_socio
    BUSCA_SOCIO
      'SET RELATION TO SOCIO INTO SOCIOS
Do Until cd.EOF = True

    With cp

    If cd.Fields("CVEMOV") = "61" Then
    'Or cd.Fields("CVEMOV") = "63" Then
        If nreg = 0 Then
            nreg = nreg + 1
           .AddNew
            cp.Fields("Id") = nreg
            cp.Fields("NOMBRE") = "FECHA DE CORTE:  " & _
                Format(PrvFecorte, "dd-mmmm-yyyy")
            cp.Update
        End If
        If cd.Fields("SOCIO") <> n_socio Then
         If nreg > 0 Then
            nreg = nreg + 1
            .AddNew
            cp.Fields("Id") = nreg
            cp.Fields("NOMBRE") = "PLAZO: " & PrvMeses & " Meses; Vence:"
            cp.Fields("FECHA") = Format(PrvFecVenc, "ddddd")
            cp.Fields("CVEMOV") = s_tasapres & "%"
            cp.Fields("IMPORTE") = Format(TotPres, "Currency")
            cp.Fields("DESCRIPCION") = "SEGURO DE SOCIOS"
            cp.Fields("PRIMA") = Format(TotPres * 0.01, "Currency")
            cp.Fields("REFERENCIA") = prvPromotor & "-" & PrvSeguro
            cp.Update
            
            nreg = nreg + 1
            .AddNew
            cp.Fields("Id") = nreg
            cp.Fields("IMPORTE") = Format(PrvPagoMin, "Currency")
            cp.Fields("DESCRIPCION") = "PAGO MINIMO"
            cp.Update
            
            nreg = nreg + 1
            .AddNew
            cp.Fields("Id") = nreg
            cp.Update
            
            TotPres = 0
            n_socio = cd.Fields("SOCIO")
         End If
        End If
    
        nreg = nreg + 1
        .AddNew
        cp.Fields("Id") = nreg
        cp.Fields("SOCIO") = cd.Fields("SOCIO")
        PrvSocio = cd.Fields("SOCIO")
        BUSCA_SOCIO     ' pubNombre = cl.Fields("NOMBRE")
        'IntRespuesta = MsgBox(pubNombre, 0)

        cp.Fields("NOMBRE") = PubNombre
        cp.Fields("FECHA") = Format(cd.Fields("FECHA"), "ddddd")
        cp.Fields("CVEMOV") = cd.Fields("CVEMOV") & "-" & cd.Fields("TIPO")
        cp.Fields("IMPORTE") = Format(cd.Fields("IMPORTE"), "Currency")
        TotPres = TotPres + cd.Fields("IMPORTE")
        cp.Fields("DESCRIPCION") = cd.Fields("DESCRIP")
        If cd.Fields("REFERENC") > "" Then
            cp.Fields("REFERENCIA") = cd.Fields("REFERENC")
        End If
        cp.Update
    End If
    End With
   cd.MoveNext
Loop
    With cp
            nreg = nreg + 1
            .AddNew
            cp.Fields("Id") = nreg
            cp.Fields("NOMBRE") = "PLAZO: " & PrvMeses & " Meses; Vence:"
            cp.Fields("FECHA") = Format(PrvFecVenc, "ddddd")
            cp.Fields("CVEMOV") = s_tasapres & "%"
            cp.Fields("IMPORTE") = Format(TotPres, "Currency")
            cp.Fields("DESCRIPCION") = "SEGURO DE SOCIOS"
            cp.Fields("PRIMA") = Format(TotPres * 0.01, "Currency")
            cp.Fields("REFERENCIA") = n_socio & "-" & PrvSeguro

            nreg = nreg + 1
            .AddNew
            cp.Fields("Id") = nreg
            cp.Fields("IMPORTE") = Format(PrvPagoMin, "Currency")
            cp.Fields("DESCRIPCION") = "PAGO MINIMO"
            cp.Update
    End With
    IntRespuesta = MsgBox("Se generó REPORTE PRESTAMOS EN DB RPRE en SYSRPT", 0)

cd.Close

End Sub

Private Sub MnuTotGrupo_Click()
    Dim cr As New ADODB.Connection
   
   Dim cd As New ADODB.Recordset

   Dim strPath As String

   cr.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisrpt.mdb"
   
    'Update the following path to point to the sample
    'Northwind.mdb database on your computer.

    strPath = "C:\" & Carpeta & "\SISRPT.mdb"

    'Create a new ADO Connection to Northwind
    'by using Access and the Jet OLE DB
    'provider.

    Set cr = New ADODB.Connection

    With cr
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

    'Create a new ADO Recordset by using a server-side
    'keyset cursor and optimistic locking.

    Set cd = New ADODB.Recordset

    With cd
        .ActiveConnection = cr
        .Source = "SELECT * FROM GRUPOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes


   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   cl.Open "SELECT * FROM SOCIOS ORDER BY GRUPO,SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   
    'Borra Registros del Archivo GRUPO
    Do Until cd.EOF = True
        cd.MoveFirst
        If Not cd.EOF Then
            cd.Delete
        End If
    Loop
    sgrupo = "00"
 'Graba Resgistros de Movimientos de Préstamos DMOVPR
             tinversion = 0
            tintganado = 0
            tcomision = 0
            tintpagado = 0
            tprestamos = 0
            tpromedio = 0
            caprestamo = 0

 cl.MoveFirst
 sgrupo = cl.Fields("GRUPO")
    Do Until cl.EOF = True
        If sgrupo <> cl.Fields("GRUPO") Then
            With cd

            PrvSocio = sgrupo
            BUSCA_SOCIO
            'IntRespuesta = MsgBox("GRUPO.-" & sgrupo & " " & pubNombre, 0)
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("GRUPO") = sgrupo
            cd.Fields("NOMBRE") = Left(PubNombre, 24)
            cd.Fields("INVERSION") = tinversion
            cd.Fields("INTGANADO") = tintganado
            cd.Fields("COMISION") = tcomision
            cd.Fields("INTPAGADO") = tintpagado
            cd.Fields("PRESTAMOS") = tprestamos
            cd.Fields("PROMEDIO") = tpromedio
            caprestamo = (tinversion + tintganado + tcomision) * 2 - tprestamos
            cd.Fields("CAPRES") = caprestamo
            cd.Fields("CAPRET") = caprestamo / 2
            
            cd.Update
        End With
            tinversion = 0
            tintganado = 0
            tcomision = 0
            tintpagado = 0
            tprestamos = 0
            tpromedio = 0
            caprestamo = 0
            sgrupo = cl.Fields("GRUPO")
        End If
        Socio = cl.Fields("SOCIO")
        tinversion = tinversion + cl.Fields("SALDO")
        tintganado = tintganado + cl.Fields("INTGANADO")
        tcomision = tcomision + cl.Fields("COMISION")
        tintpagado = tintpagado + cl.Fields("INTPAGADO")
        tprestamos = tprestamos + cl.Fields("SALDOPRES")
        If cl.Fields("PROM_INV") <> "" Then
            tpromedio = tpromedio + cl.Fields("PROM_INV")
        End If
        'If sgrupo = "02" Then
        '    IntRespuesta = MsgBox("SOCIO=" & sgrupo & " " & Socio & " INTGANADO=" & cl.Fields("INTGANADO") & " " & tintganado, 0)
        'End If

    cl.MoveNext
Loop
    
   IntRespuesta = MsgBox("Se generó REPORTE DE GRUPOS PARA EXCEL en DB GRUPOS en SYSRPT", 0)


End Sub

Private Sub SaldosDBF_Click()
IntRespuesta = MsgBox("¡¡¡CUIDADO!!!       ESTE PROCESO COPIA SALDOS DBF", 1)
If (IntRespuesta = 1) Then
    IntRespuesta = MsgBox("Continúa", 0)
Else
    Exit Sub
End If
IntRespuesta = MsgBox("¿ESTA SEGURO DE COPIAR SALDOS DBF...?", 1)
If (IntRespuesta = 1) Then
    IntRespuesta = MsgBox("Continúa", 0)
Else
    Exit Sub
End If

    'COPIA DATOS de SOCIOS1 a SOCIOS
    
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el Objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM SOCIOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    'Inicializa DATOS del Archivo SOCIOS
    cl.MoveFirst

    Do Until cl.EOF = True
        PrvSocio = cl.Fields("SOCIO")
        BUSCA_SOCIO_DBF
        cl.Fields("SALDO") = Saldo_DBF
        cl.Fields("INTGANADO") = IntGanado_DBF
        cl.Fields("COMISION") = Comision_DBF
        cl.Fields("SALDOPRES") = SaldoPres_DBF
        cl.Update
        cl.MoveNext
    Loop
   
    
    IntRespuesta = MsgBox("TERMINA COPIA SALDOS DBF", 0)

End Sub

Private Sub RelGrupos_Click()
    Dim cr As New ADODB.Connection
   
   Dim cd As New ADODB.Recordset

   Dim strPath As String

   cr.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisrpt.mdb"
   
    'Update the following path to point to the sample
    'Northwind.mdb database on your computer.

    strPath = "C:\" & Carpeta & "\SISRPT.mdb"

    'Create a new ADO Connection to Northwind
    'by using Access and the Jet OLE DB
    'provider.

    Set cr = New ADODB.Connection

    With cr
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

    'Create a new ADO Recordset by using a server-side
    'keyset cursor and optimistic locking.

    Set cd = New ADODB.Recordset

    With cd
        .ActiveConnection = cr
        .Source = "SELECT * FROM RGRP"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes


   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   cl.Open "SELECT * FROM SOCIOS ORDER BY GRUPO, SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   
    Do Until cd.EOF = True
        cd.MoveFirst
        If Not cd.EOF Then
            cd.Delete
        End If
    Loop
    sgrupo = "00"
    tinversion = 0
            tintganado = 0
            tcomision = 0
            tintpagado = 0
            tprestamos = 0
            tpromedio = 0
            caprestamo = 0
            PorcienGrupo = 0

 cl.MoveFirst
 sgrupo = cl.Fields("GRUPO")
    Do Until cl.EOF = True
        With cd

        If sgrupo <> cl.Fields("GRUPO") Then

            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            If tprestamos > 0 Then
                PorcienGrupo = (tprestamos / (tinversion + tintpagado + tcomision)) * 100
            End If
           ' MsgBox ("Se generó=" & PorcienGrupo & "%")
            If PorcienGrupo > 0 Then
                cd.Fields("SOCIO") = PorcienGrupo & "%"
            End If
            cd.Fields("NOMBRE") = "Préstamos VS Inversión"
            cd.Fields("INVERSION") = tinversion
            cd.Fields("INTGANADO") = tintganado
            cd.Fields("COMISION") = tcomision
            cd.Fields("INTPAGADO") = tintpagado
            cd.Fields("PRESTAMOS") = tprestamos
            cd.Fields("PROMEDIO") = tpromedio
            
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("NOMBRE") = "Capacidad de Préstamos"
            caprestamo = (tinversion + tintganado + tcomision) * 2 - tprestamos
            cd.Fields("INVERSION") = caprestamo
            
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("NOMBRE") = "Capacidad de Retiros"
            cd.Fields("INVERSION") = caprestamo / 2
            
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Update
            
            sgrupo = cl.Fields("GRUPO")
            tinversion = 0
            tintganado = 0
            tcomision = 0
            tintpagado = 0
            tprestamos = 0
            tpromedio = 0
            caprestamo = 0
            PorcienGrupo = 0
        End If
        
            PrvSocio = cl.Fields("SOCIO")
            BUSCA_SOCIO
            nreg = nreg + 1
            .AddNew
            cd.Fields("Id") = nreg
            cd.Fields("GRUPO") = cl.Fields("GRUPO")
            PubNombre = cl.Fields("NOMBRE")
            cd.Fields("NOMBRE") = Left(PubNombre, 24)
            cd.Fields("SOCIO") = cl.Fields("SOCIO")
            cd.Fields("INVERSION") = cl.Fields("SALDO")
            cd.Fields("INTGANADO") = cl.Fields("INTGANADO")
            cd.Fields("COMISION") = cl.Fields("COMISION")
            cd.Fields("INTPAGADO") = cl.Fields("INTPAGADO")
            cd.Fields("PRESTAMOS") = cl.Fields("SALDOPRES")
            cd.Fields("PROMEDIO") = cl.Fields("PROM_INV")
            tinversion = tinversion + cl.Fields("SALDO")
            tintganado = tintganado + cl.Fields("INTGANADO")
            tcomision = tcomision + cl.Fields("COMISION")
            tintpagado = tintpagado + cl.Fields("INTPAGADO")
            tprestamos = tprestamos + cl.Fields("SALDOPRES")
            If cl.Fields("PROM_INV") <> "" Then
                tpromedio = tpromedio + cl.Fields("PROM_INV")
            End If
            
            cd.Update
            Socio = cl.Fields("SOCIO")
    End With
    cl.MoveNext
Loop
   
   IntRespuesta = MsgBox("Se generó RELACION DE GRUPOS PARA EXCEL en DB RGRP en SYSRPT", 0)

End Sub

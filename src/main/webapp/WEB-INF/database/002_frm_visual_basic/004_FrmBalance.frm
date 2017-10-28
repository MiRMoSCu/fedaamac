VERSION 5.00
Begin VB.Form FrmBalance 
   Caption         =   "BALANCE"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtDepPagos 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   65
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox TxtDepBanca 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   63
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox TxtBalMRL 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   61
      Text            =   "BANCO MRL"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox TxtBalMLB 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   59
      Text            =   "BCO MLB"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox TxtTotBancos 
      Alignment       =   1  'Right Justify
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
      Left            =   2520
      TabIndex        =   56
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox TxtBalFecorte 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox TxtBalLLB 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   55
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox TxtBalGNL 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   54
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox TxtBalVCL 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   53
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox TxtBalJBS 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   52
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox TxtBalGMF 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   51
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox TxtBalBMX 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   50
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox TxtBalTotBancos 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   49
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox TxtBalIntInv 
      Alignment       =   1  'Right Justify
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
      Left            =   7440
      TabIndex        =   48
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox TxtBalTasaAnual 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   47
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox TxtBalTasa 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
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
      Left            =   7440
      TabIndex        =   46
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox TxtBalPrestamos 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   45
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox TxtBalInvBanca 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   44
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox TxtBalISR 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   43
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox TxtBalDif 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   38
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox TxtBalAportaciones 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   36
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox TxtBalCapital 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   34
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox TxtBalIntGanado 
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
      Left            =   7440
      TabIndex        =   32
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox TxtBalSeguro 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   30
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox TxtBalComPres 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   28
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox TxtBalReserva 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   26
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox TxtBalIntBco 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   24
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox TxtBalInvMor 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   22
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox TxtBalAporMor 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   21
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox TxtBalPasivo 
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
      Height          =   285
      Left            =   7440
      TabIndex        =   19
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox TxtBalActivo 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox TxtBalRecup 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   16
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label34 
      Caption         =   "DEPOSITOS POR PAGOS"
      Height          =   255
      Left            =   240
      TabIndex        =   64
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label33 
      Caption         =   "DEPOSITOS POR AHORRO"
      Height          =   255
      Left            =   240
      TabIndex        =   62
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label32 
      Caption         =   " "
      Height          =   255
      Left            =   240
      TabIndex        =   60
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Lbl991 
      Caption         =   "BANCO MLB"
      Height          =   255
      Left            =   240
      TabIndex        =   58
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label30 
      Caption         =   "TOTAL BANCOS"
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label29 
      Caption         =   "LOS INTERESES S/INVERSION ESTAN INTEGRADOS EN LOS INTERESES GANADOS POR LOS SOCIOS"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   7320
      Width           =   8895
   End
   Begin VB.Label Label28 
      Caption         =   "INTERESES S/INVERSION     "
      Height          =   255
      Left            =   5040
      TabIndex        =   41
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label27 
      Caption         =   "TASA ANUAL                    "
      Height          =   255
      Left            =   5040
      TabIndex        =   40
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label26 
      Caption         =   "TASA DE INVERSION             "
      Height          =   255
      Left            =   5040
      TabIndex        =   39
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label25 
      Caption         =   "DIFERENCIA           "
      Height          =   255
      Left            =   4920
      TabIndex        =   37
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label24 
      Caption         =   "APORTACIONES        "
      Height          =   375
      Left            =   5040
      TabIndex        =   35
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label23 
      Caption         =   "CAPITAL SOCIAL"
      Height          =   255
      Left            =   5040
      TabIndex        =   33
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label22 
      Caption         =   "INTERESES GANADOS        "
      Height          =   255
      Left            =   5040
      TabIndex        =   31
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label21 
      Caption         =   "SEGURO DE SOCIOS         "
      Height          =   255
      Left            =   5040
      TabIndex        =   29
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label20 
      Caption         =   "COMISION X PRESTAMOS     "
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label19 
      Caption         =   "RESERVA PARA PREMIOS     "
      Height          =   255
      Left            =   5040
      TabIndex        =   25
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label18 
      Caption         =   "INTERESES DEL BANCO"
      Height          =   255
      Left            =   5040
      TabIndex        =   23
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label17 
      Caption         =   "APORTACIONES MOROLEON"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label16 
      Caption         =   "TOTAL PASIVO"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "RECUPERACION DEL MES"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label14 
      Caption         =   "TOTAL ACTIVO"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "INVERSION MOROLEON"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "ISR Y COMISIONES"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "INVERSION BANCARIA"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "PRESTAMOS"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "TOTAL DISPONIBE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Lbl993 
      Caption         =   "BANCO BMX"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Lbl992 
      Caption         =   "BANCO GMF"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Lbl994 
      Caption         =   "BANCO JBS"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Lbl995 
      Caption         =   "BANCO VCL"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Lbl996 
      Caption         =   "BANCO GNL"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Lbl990 
      Caption         =   "99  .-BANCO LLB HSBC"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "BALANCE GENERAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "FONDO ECONOMICO DE AHORRO Y AYUDA MUTUA, A.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   11295
   End
End
Attribute VB_Name = "FrmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Carpeta As String
Private Sub Form_Load()
'  IntRespuesta = MsgBox("Mi Primera Carpeta=" & frmMiPrimera.LblCarpeta, 0)
    Carpeta = frmMiPrimera.LblCarpeta
'    IntRespuesta = MsgBox("Carpeta=" & Carpeta, 0)
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb" '''
    TxtBalance
End Sub



Private Sub Label31_Click()

End Sub

Private Sub TxtBalFecorte_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Or KeyAscii = 13 Then 'vbKeyesc vbKeyEnter
        'IntRespuesta = MsgBox("KeyAscii=" & KeyAscii & "-" & prvImporte, 0)
        'lfrmCount = lfrmCount - 1
        'Set frmD = frmMENUSYS
    'frmD.Caption = "frmMENUSYS"
    
    'frmD.Show
    TxtBalance
    End If
End Sub
Private Sub TxtBalFecorte_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub TxtBalFecorte_Change()
    TxtBalance
End Sub
  '    * ³              Nombre: R_BALANCE            Document¢: Luis Lopez Baeza            ³
'* ³         Descripci¢n: BALANCE DE LA CUENTA DE FEDAMAC AL ULTIMO CORTE             ³
'* ³               Autor: Luis L¢pez Baeza                                            ³
'* ³   Fecha de creaci¢n: 09-06-2010            Fecha de Actualizaci¢n:               ³
'* ³    Hora de creaci¢n: 6:21 pm               Hora de Actualizaci¢n:                ³
'* ³ Derechos Reservados: LOBA FEDAMAC S.A. DE C.V.                                   ³
'* ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
'* ³        Argumentos: None                                                          ³
'* ³ Valor que Regresa: Nil                                                           ³
'* ³       Ver Tambi‚n:                                                               ³
'* ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
'Function r_balance()
Private Sub TxtBalance()
    Dim s_aporta, s_prestamo, s_activo, s_pasivo, s_intganado, s_vencido As Double
    Dim s_comision, s_comban, s_reserva, s_intbanco, s_actfijo, s_deprecia As Double
    Dim s_recup, s_seguro, s_totini, s_invban, s_intinver, s_bcobmx As Double
    Dim s_caja, s_bcogmf As Double
    
    Dim m_socio, m_socio1, m_socio2 As String
    
    '* Abrimos bases de datos
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SOCIOS ORDER BY TIPO, SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
    TxtBalFecorte.Text = cl.Fields("FECORTE")
    s_corte = cl.Fields("FECORTE")
    's_tasainv = cl.Fields("INTGANADO") / cl.Fields("PROM_INV") * 100
    m_socio1 = "001"
    m_socio2 = "999"
    s_actual = 0
    s_aporta = 0
                     'IntRespuesta = MsgBox(cl.Fields("SOCIO"), 0)

    Do Until cl.EOF = True
         If cl.Fields("SOCIO") = 25 Then
            s_intbanco = s_intbanco + cl.Fields("SALDO")
            s_comban = s_comban + cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            cl.MoveNext
         End If

         If cl.Fields("SOCIO") = 48 Then
            s_morfin = cl.Fields("SALDO")
            s_invmor = cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            cl.MoveNext
            
         End If
         
         If cl.Fields("SOCIO") = 99 Then
            s_caja = cl.Fields("SALDO")
            s_activo = s_activo + cl.Fields("SALDO")
            cl.MoveNext
         End If
                  
         If cl.Fields("SOCIO") = 990 Then
            s_bcomrl = cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            cl.MoveNext
         End If

         If cl.Fields("SOCIO") = 991 Then
            sNombre = cl.Fields("NOMBRE")
            Lbl991 = ("991.-" & sNombre)
            s_bcomlb = cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            cl.MoveNext
         End If

         If cl.Fields("SOCIO") = 992 Then
            sNombre = cl.Fields("NOMBRE")
            Lbl992 = ("992.-" & sNombre)
            s_bcogmf = cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            cl.MoveNext
         End If

         If cl.Fields("SOCIO") = 993 Then
            sNombre = cl.Fields("NOMBRE")
            Lbl993 = ("993.-" & sNombre)
            s_bcobmx = cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            cl.MoveNext
         End If
         
         If cl.Fields("SOCIO") = 994 Then
            sNombre = cl.Fields("NOMBRE")
            Lbl994 = ("994.-" & sNombre)
            s_bcomrr = cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            cl.MoveNext
         End If
         
         If cl.Fields("SOCIO") = 995 Then
            sNombre = cl.Fields("NOMBRE")
            Lbl995 = ("995.-" & sNombre)
            s_bcocvl = cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            cl.MoveNext
         
         End If
         
         If cl.Fields("SOCIO") = 996 Then
            sNombre = cl.Fields("NOMBRE")
            Lbl996 = ("996.-" & sNombre)
            s_bcognl = cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            cl.MoveNext
         End If
         
         If cl.Fields("SOCIO") = "988" Then
            s_invban = s_invban + cl.Fields("SALDOPRES")
            s_activo = s_activo + cl.Fields("SALDOPRES")
            s_intinver = s_intinver + cl.Fields("SALDO")
            's_intganado = s_intganado + cl.Fields("SALDO")
            s_intganado = s_intganado + cl.Fields("INTGANADO")
            cl.MoveNext
         End If
         
         If cl.Fields("TIPO") = 8 Then
            s_seguro = s_seguro + cl.Fields("SALDO")
            s_aporta = s_aporta - cl.Fields("SALDO")
            's_intganado = s_intganado + cl.Fields("INTGANADO")
            'cl.MoveNext
         End If

         If cl.Fields("SOCIO") = 50 Then
            s_tasainv = cl.Fields("INTGANADO") / cl.Fields("PROM_INV") * 100
            s_reserva = cl.Fields("COMISION")
            s_comision = s_comision - cl.Fields("COMISION")
            s_capsal = cl.Fields("SALDO")
            s_aporta = s_aporta - cl.Fields("SALDO")
         End If
         
         s_aporta = s_aporta + cl.Fields("SALDO")
         s_prestamos = s_prestamos + cl.Fields("SALDOPRES")
         s_intganado = s_intganado + cl.Fields("INTGANADO")
         s_activo = s_activo + cl.Fields("SALDOPRES")
         s_comision = s_comision + cl.Fields("COMISION")
         If cl.Fields("SALDOPRES") > 0 Then
            s_recup = s_recup + cl.Fields("PAGOMIN")
         End If
       cl.MoveNext
    Loop
cl.Close

        t_bancos = s_caja + s_bcognl + s_bcocvl + s_bcomrr + s_bcobmx + s_bcogmf + s_bcomlb + s_bcomrl
        s_pasivo = s_aporta + s_intganado + s_comision + s_reserva + s_morfin + s_seguro + s_capsal
                'MS.Text = Format(cl.Fields("SALDOPRES"), "Currency")

        TxtBalLLB = Format(s_caja, "Currency")
        TxtBalAportaciones = Format(s_aporta, "Currency")
        TxtBalGNL = Format(s_bcognl, "Currency")
        TxtBalCapital = Format(s_capsal, "Currency")
        TxtBalVCL = Format(s_bcocvl, "Currency")
        TxtBalJBS = Format(s_bcomrr, "Currency")
        TxtBalGMF = Format(s_bcogmf, "Currency")
        TxtBalMLB = Format(s_bcomlb, "Currency")
        TxtBalMRL = Format(s_bcomrl, "Currency")
        TxtBalBMX = Format(s_bcobmx, "Currency")
        TxtTotBancos = Format(t_bancos, "Currency")
        TxtBalTotBancos = Format(t_bancos + s_invban, "Currency")
        TxtBalPrestamos = Format(s_prestamos, "Currency")
        TxtBalIntGanado = Format(s_intganado, "Currency")
        TxtBalInvBanca = Format(s_invban, "Currency")
        TxtBalSeguro = Format(s_seguro, "Currency")
        TxtBalComPres = Format(s_comision, "Currency")
        TxtBalReserva = Format(s_reserva, "Currency")
        TxtBalISR = Format(s_comban, "Currency")
        TxtBalIntBanco = Format(s_intbanco, "Currency")
        TxtBalInvMor = Format(s_invmor, "Currency")
        TxtBalAporMor = Format(s_morfin, "Currency")
        TxtBalActivo = Format(s_activo, "Currency")
        TxtBalPasivo = Format(s_pasivo, "Currency")
        TxtBalRecup = Format(s_recup, "Currency")
        TxtBalDif = Format(s_activo - s_pasivo, "Currency")
        TxtBalTasa = Format(s_tasainv / 100, "Percent")
        
        DepBanca

        If Month(s_corte) < 11 Then
            s_numes = Month(s_corte) + 2
        Else
            s_numes = Month(s_corte) - 10
        End If
        TxtBalTasaAnual = Format(s_tasainv / s_numes * 12 / 100, "Percent")
        TxtBalIntInv = Format(s_intinver, "Currency")

End Sub
Private Sub DepBanca()
    Dim s_DepBanca As Double
    Dim s_DepPagos As Double
    Dim s_fecha As Date
    
    '* Abrimos bases de datos
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SICMOV ORDER BY TIPO, APREPAC", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
    s_DepBanca = 0
    s_DepPagos = 0
                   'IntRespuesta = MsgBox(cl.Fields("SOCIO"), 0)

    Do Until cl.EOF = True
        s_fecha = cl.Fields("FECHA")
    If cl.Fields("CVEMOV") <> "59" Then
        If Month(s_fecha) = Month(TxtBalFecorte) Then
         If cl.Fields("TIPO") = "B" Then
            If cl.Fields("APREPAC") = "A" Or cl.Fields("APREPAC") = "P" Then
                'IntRespuesta = MsgBox(cl.Fields("FECHA") & " " & cl.Fields("TIPO") & " " & cl.Fields("APREPAC") & " " & cl.Fields("IMPORTE") & " " & s_DepBanca, 0)
                If cl.Fields("CVEMOV") = "10" Or cl.Fields("CVEMOV") = "11" Or cl.Fields("CVEMOV") = "12" Then
                    s_DepBanca = s_DepBanca + cl.Fields("IMPORTE")
                    'IntRespuesta = MsgBox(cl.Fields("FECHA") & " " & cl.Fields("TIPO") & " " & cl.Fields("APREPAC") & " " & cl.Fields("IMPORTE") & " " & s_DepBanca, 0)
                End If
                    If cl.Fields("CVEMOV") = "50" Or cl.Fields("CVEMOV") = "51" Or cl.Fields("CVEMOV") = "52" Then
                    s_DepPagos = s_DepPagos + cl.Fields("IMPORTE")
                    'IntRespuesta = MsgBox(cl.Fields("FECHA") & " " & cl.Fields("TIPO") & " " & cl.Fields("APREPAC") & " " & cl.Fields("IMPORTE") & " " & s_DepBanca, 0)
                End If

            End If
         End If
        End If
    End If
        cl.MoveNext
        
        Loop
   TxtDepBanca = Format(s_DepBanca, "Currency")
   TxtDepPagos = Format(s_DepPagos, "Currency")
End Sub

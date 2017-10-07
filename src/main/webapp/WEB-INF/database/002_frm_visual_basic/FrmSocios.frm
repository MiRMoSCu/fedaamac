VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSocios 
   BackColor       =   &H8000000E&
   Caption         =   "2"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   14760
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   100
      Text            =   "."
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox PicSocio 
      Height          =   2055
      Left            =   12120
      Picture         =   "FrmSocios.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   1755
      TabIndex        =   98
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox TxtNumSocio 
      Height          =   285
      Left            =   12120
      TabIndex        =   1
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton CmdMovpr 
      Caption         =   "Consulta Movimientos de Préstamos"
      Height          =   735
      Left            =   12360
      TabIndex        =   4
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton CmdMovin 
      Caption         =   "Consulta Movimientos de Inversion"
      Height          =   735
      Left            =   12360
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox TxtNomSocio 
      Height          =   285
      Left            =   12360
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   0
      TabIndex        =   95
      Top             =   0
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid FGS 
      Height          =   5520
      Left            =   240
      TabIndex        =   94
      Top             =   3720
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9737
      _Version        =   393216
      Rows            =   2000
      Cols            =   12
   End
   Begin VB.TextBox TxtUltPago 
      Height          =   285
      Left            =   12240
      TabIndex        =   93
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox TxtSeguro 
      Height          =   285
      Left            =   6000
      TabIndex        =   21
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton CmdEctaPre 
      Caption         =   "IMPRESION ESTADO DE CUENTA DE PRESTAMOS"
      Height          =   615
      Left            =   10800
      TabIndex        =   89
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton CmdEdoCta 
      Caption         =   "IMPRESION ESTADO DE CUENTA DE INVERSION"
      Height          =   615
      Left            =   8280
      TabIndex        =   88
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton CmdActualizaSocio 
      Caption         =   "ACTUALIZA SOCIO"
      Height          =   495
      Left            =   6480
      TabIndex        =   87
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton CmdBorraReg 
      Caption         =   "ELIMINA SOCIO"
      Height          =   495
      Left            =   5040
      TabIndex        =   86
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox TxtclPromotor 
      Height          =   285
      Left            =   4800
      TabIndex        =   20
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox TxtclTipo 
      Height          =   285
      Left            =   3600
      TabIndex        =   19
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox TxtclGrupo 
      Height          =   285
      Left            =   2640
      TabIndex        =   18
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton CmdNvoSocio 
      Caption         =   "ALTA NUEVO SOCIO"
      Height          =   495
      Left            =   5040
      TabIndex        =   82
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox TxtTasaPres 
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
      Height          =   285
      Left            =   9480
      TabIndex        =   80
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox TxtSaldoCaja 
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
      Left            =   12240
      TabIndex        =   78
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox TxtFPrestamo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   12240
      TabIndex        =   76
      Text            =   "31/08/1989"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox TxtFVencimiento 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   12240
      TabIndex        =   75
      Text            =   "31/08/1989"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox TxtInvMoroleon 
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
      Height          =   285
      Left            =   12240
      TabIndex        =   74
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox TxtTasaInversion 
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
      Height          =   285
      Left            =   12240
      MaxLength       =   5
      TabIndex        =   73
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox TxtPromInversion 
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
      Height          =   285
      Left            =   12240
      TabIndex        =   72
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox TxtPromAportacion 
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
      Height          =   285
      Left            =   12240
      TabIndex        =   71
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox TxtFechaCorte 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   12240
      TabIndex        =   70
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox TxtPagoTotal 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   62
      Top             =   2400
      Width           =   1335
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
      Height          =   285
      Left            =   9360
      TabIndex        =   61
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtPrestamoActual 
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
      Left            =   9360
      TabIndex        =   60
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox TxtIntsPagados 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9360
      TabIndex        =   59
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox TxtPagos 
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
      Height          =   285
      Left            =   9360
      TabIndex        =   58
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox TxtPrestamos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9360
      TabIndex        =   57
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtPrestamoInicial 
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
      Height          =   285
      Left            =   9360
      TabIndex        =   56
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox TxtComisiones 
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
      Height          =   285
      Left            =   6480
      TabIndex        =   47
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox TxtIntsDevengados 
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
      Height          =   285
      Left            =   6480
      TabIndex        =   46
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox TxtSaldoActual 
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
      Left            =   6480
      TabIndex        =   45
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox TxtRetiros 
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
      Height          =   285
      Left            =   6480
      TabIndex        =   44
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox TxtAportaciones 
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
      Height          =   285
      Left            =   6480
      TabIndex        =   43
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox TxtSaldoInicial 
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
      Left            =   6480
      TabIndex        =   42
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TxtclApertura 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   32
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtclFecNac 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   31
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox TxtclCLABE 
      Height          =   285
      Left            =   1560
      TabIndex        =   30
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox TxtcleMail 
      Height          =   285
      Left            =   1560
      TabIndex        =   29
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox TxtclTelefono 
      Height          =   285
      Left            =   1560
      TabIndex        =   28
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TxtclCP 
      Height          =   285
      Left            =   1560
      TabIndex        =   27
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox TxtclEstado 
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   26
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox TxtclCiudad 
      Height          =   285
      Left            =   1560
      TabIndex        =   25
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox TxtclColonia 
      Height          =   285
      Left            =   1560
      TabIndex        =   24
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox TxtclDireccion 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   23
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox TxtclNombre 
      Height          =   285
      Left            =   1560
      TabIndex        =   22
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox TxtclSocio 
      Height          =   285
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   17
      Text            =   "01"
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label49 
      Caption         =   "Asistencia"
      Height          =   255
      Left            =   360
      TabIndex        =   99
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label48 
      Caption         =   "Lista Num Socios"
      Height          =   255
      Left            =   12120
      TabIndex        =   97
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label47 
      Caption         =   "Nombre del Socio"
      Height          =   255
      Left            =   12360
      TabIndex        =   96
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label46 
      Caption         =   "Ultimo Pago"
      Height          =   255
      Left            =   10800
      TabIndex        =   92
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label45 
      Caption         =   "Seguro"
      Height          =   255
      Left            =   5280
      TabIndex        =   91
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label44 
      Caption         =   "%"
      Height          =   255
      Left            =   10440
      TabIndex        =   90
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label43 
      Caption         =   "Promotor"
      Height          =   255
      Left            =   3960
      TabIndex        =   85
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label42 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   3120
      TabIndex        =   84
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label41 
      Caption         =   "Grupo"
      Height          =   255
      Left            =   2040
      TabIndex        =   83
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   13800
      TabIndex        =   81
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label39 
      Caption         =   "Tasa Préstamo"
      Height          =   255
      Left            =   8280
      TabIndex        =   79
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label38 
      Caption         =   "Saldo en Caja"
      Height          =   255
      Left            =   10800
      TabIndex        =   77
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label37 
      Caption         =   "F. Prestamo  "
      Height          =   255
      Left            =   10800
      TabIndex        =   69
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label36 
      Caption         =   "F. Vencimiento  "
      Height          =   255
      Left            =   10800
      TabIndex        =   68
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label35 
      Caption         =   "Inv. Moroleón"
      Height          =   255
      Left            =   10800
      TabIndex        =   67
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "Tasa de Inversión"
      Height          =   255
      Left            =   10800
      TabIndex        =   66
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label33 
      Caption         =   "Prom. Inversión"
      Height          =   255
      Left            =   10800
      TabIndex        =   65
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label32 
      Caption         =   "Fecha de Corte"
      Height          =   255
      Left            =   10800
      TabIndex        =   64
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label31 
      Caption         =   "Prom. Aportación"
      Height          =   255
      Left            =   10800
      TabIndex        =   63
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label30 
      Caption         =   "Pago Total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   55
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label29 
      Caption         =   "Pago Minimo        "
      Height          =   255
      Left            =   8280
      TabIndex        =   54
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "Saldo Actual       "
      Height          =   255
      Left            =   8280
      TabIndex        =   53
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label27 
      Caption         =   "Ints. Pagados  "
      Height          =   255
      Left            =   8280
      TabIndex        =   52
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label26 
      Caption         =   "Pagos "
      Height          =   255
      Left            =   8280
      TabIndex        =   51
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "Préstamos "
      Height          =   255
      Left            =   8280
      TabIndex        =   50
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "Saldo Inicial     "
      Height          =   255
      Left            =   8280
      TabIndex        =   49
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "RESUMEN DE PRESTAMOS"
      Height          =   255
      Left            =   8280
      TabIndex        =   48
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label22 
      Caption         =   "Comisiones"
      Height          =   255
      Left            =   5160
      TabIndex        =   41
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Ints. Devengados          "
      Height          =   255
      Left            =   5160
      TabIndex        =   40
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "Saldo Actual             "
      Height          =   255
      Left            =   5160
      TabIndex        =   39
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "Retiros  "
      Height          =   255
      Left            =   5160
      TabIndex        =   38
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label18 
      Caption         =   "Aportaciones"
      Height          =   255
      Left            =   5160
      TabIndex        =   37
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "Saldo Inicial       "
      Height          =   255
      Left            =   5160
      TabIndex        =   36
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   135
      Left            =   7800
      TabIndex        =   35
      Top             =   960
      Width           =   15
   End
   Begin VB.Label Label15 
      Caption         =   "RESUMEN DE INVERSION  "
      Height          =   255
      Left            =   5640
      TabIndex        =   34
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "CONSULTA  DE SOCIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   33
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label13 
      Caption         =   "Fecha de Apertura"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha de Nacimiento"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "CLABE Cta Bco"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "eMail"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Código Postal"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Estado"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Ciudad"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Colonia"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Socio"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Public Carpeta As String
    Public PubNombre As String
    Private PrvSocio As String
    Private RenBorra, Celda As Single
    Private CLABE As String
    Private PrvCita1, PrvCita2, PrvCita3 As String
    Private flgsocio As Integer
    
Sub Busca_Cita()
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM CITAS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cd.MoveFirst
    Do Until cd.EOF = True
        totreg = totreg + 1
        cd.MoveNext
        Loop
    cd.MoveFirst
    
    Randomize
    Aleatorio = CLng((1 - totreg) * Rnd + totreg)

    Do Until cd.EOF = True
        If Aleatorio = numreg Then
            PrvCita1 = cd.Fields("CITA1")
            If cd.Fields("CITA2") > "" Then
                PrvCita2 = cd.Fields("CITA2")
            Else
                PrvCita2 = ""
            End If
            If cd.Fields("CITA3") > "" Then
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


Private Sub Command1_Click()
    Dim BD As Database
    Dim RSSOCIOS As Recordset
    Dim Ruta As String
    Ruta = "c:\" & Carpeta & "\sisfed.mdb"
    Set BD = OpenDatabase(Ruta)
    setRSSOCIOS = BD.OpenRecordset(SQL, dbOpenDynaset)
    'select * from socios"
    RSSOCIOS.AddNew     '(SI SE QUIERE ACTUALIZAR SOLO CAMBIA A: RSSOCIOS.EDIT)
    RSSOCIOS.Fields("numero") = 200                        'convierte en entero un campo con contenido numerico
    RSSOCIOS.Fields("NOMBRE") = "Nombre"       'ELIMINA BLANCOS  A LA DERECHA
    RSSOCIOS.Fields("saldo") = 0            'CONVIERTE EN  CAMPO CURRENCY (MONEDA)
    RSSOCIOS.Update
    RSSOCIOS.Close

End Sub

Private Sub CmdActualizaSocio_Click()
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
        If cl.Fields("SOCIO") = TxtclSocio Then
            PrvSocio = TxtclSocio
        cl.Fields("GRUPO") = TxtclGrupo
        cl.Fields("SOCIO") = TxtclSocio
        cl.Fields("TIPO") = TxtclTipo
        cl.Fields("PROMOTOR") = TxtclPromotor
        cl.Fields("NOMBRE") = TxtclNombre
        cl.Fields("DIRECCION") = TxtclDireccion
        cl.Fields("COLONIA") = TxtclColonia
        cl.Fields("CIUDAD") = TxtclCiudad
        cl.Fields("ESTADO") = TxtclEstado
        cl.Fields("CP") = TxtclCP
        cl.Fields("TELEFONO") = TxtclTelefono
        cl.Fields("EMAIL") = TxtcleMail
        cl.Fields("CLABE") = TxtclCLABE
        cl.Fields("FECNAC") = txtclFecNac
        cl.Fields("FECAPER") = TxtclApertura
        cl.Fields("FECVENC") = TxtFVencimiento
        cl.Fields("FECPRES") = TxtFPrestamo
        cl.Fields("CTASEGURO") = TxtSeguro
        cl.Fields("PRES_INI") = TxtPrestamoInicial
        If TxtTasaPres <> "" Then
            cl.Fields("TASAPRES") = TxtTasaPres
        End If
        If TxtPagoMin <> "" Then
            cl.Fields("PAGOMIN") = TxtPagoMin
        End If
        cl.Update
        IntRespuesta = MsgBox("El registro se actualizó correctamente", 0)
        TxtclSocio = PrvSocio
        Exit Do
    End If
    cl.MoveNext
        
    Loop
    'TxtclSocio = ""
End Sub

Private Sub CmdBorraReg_Click()
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
        If cl.Fields("SOCIO") = TxtclSocio Then
            Exit Do
        Else
            cl.MoveNext
        End If
        Loop
    If Not cl.EOF Then
    
        IntRespuesta = MsgBox("El socio No. " & " " & cl.Fields("SOCIO") & " Será borrado", 1)
        If (IntRespuesta = 1) Then

            If cl.Fields("SALDO") < 0.01 And cl.Fields("RETIROS") < 0.01 And cl.Fields("SALDOPRES") < 0.01 And cl.Fields("PAGOS") < 0.01 Then
                cl.Delete
                TxtclSocio = " "
                TxtclNombre = " "
            Else
                IntRespuesta = MsgBox("El socio NO PUEDE SER BORRADO", 0)
            
            End If
        End If
    End If
End Sub

Private Sub CmdEctaPre_Click()

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim cp As New ADODB.Recordset 'Creamos el objeto Recordset.DMOVPR

   Dim strPath As String

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   cp.Open "SELECT * FROM DMOVPR ORDER BY SOCIO,FECHA,APREPAC DESC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cl.MoveFirst

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
        Word.Selection.TypeText "                            ESTADO DE CUENTA" & vbCrLf
        Word.Selection.TypeText "-----------------------------------------------------------------------------------" & vbCrLf
        Word.Selection.TypeText "Socio.-" & PrvSocio & ".-"
        LONGITUD = Len(PubNombre)
        LONGITUD = 48 - LONGITUD
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

Sub COLOCATITULOSENSOCIOS()
Numovs = 2000
FGS.Row = 0
FGS.Col = 0
FGS.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
FGS.Col = 1
FGS.Text = "SOCIO"
FGS.Col = 2
FGS.Text = "FECHA"
FGS.Col = 3
FGS.Text = "CVE"
FGS.Col = 4
FGS.Text = "T"
FGS.Col = 5
FGS.Text = "C"
FGS.Col = 6
FGS.Text = "DESCRIPCION"
FGS.Col = 7
FGS.Text = "REFERENCIA"
FGS.Col = 8
FGS.Text = "BANCO"
FGS.Col = 9
FGS.Text = "DEPOSITOS"
FGS.Col = 10
FGS.Text = "RETIROS"
FGS.Col = 11
FGS.Text = "SALDO"
                    'AJUSTO EL ANCHO DE LAS COLUMNAS
FGS.ColWidth(0) = 450
               
FGS.ColWidth(1) = 600
FGS.ColWidth(2) = 1000

FGS.ColWidth(3) = 500
FGS.ColWidth(4) = 200
FGS.ColWidth(5) = 200

FGS.ColWidth(6) = 2350
FGS.ColWidth(7) = 1150

FGS.ColWidth(8) = 650
FGS.ColWidth(9) = 1100
FGS.ColWidth(10) = 1100

FGS.ColWidth(11) = 1300

sdoActual = 0
End Sub
Sub BorraCeldasenSOCIOS()
    RENGLON = RenBorra
    Do Until RENGLON = 1999
    
       RENGLON = RENGLON + 1
       FGS.Col = 0
       FGS.Row = RENGLON
       FGS.Text = ""
       FGS.Col = 1
       FGS.Text = ""
       FGS.Col = 2
       FGS.Text = ""
       FGS.Col = 3
       FGS.Text = ""
       FGS.Col = 4
       FGS.Text = ""
       FGS.Col = 5
       FGS.Text = ""
       FGS.Col = 6
       FGS.Text = ""
       FGS.Col = 7
       FGS.Text = ""
       FGS.Col = 8
       FGS.Text = ""
       FGS.Col = 9
       FGS.Text = ""
       FGS.Col = 10
       FGS.Text = ""
       FGS.Col = 11
       FGS.Text = ""

    Loop

End Sub

Private Sub CmdEdoCta_Click()

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim cp As New ADODB.Recordset 'Creamos el objeto Recordset.DMOVIN

   Dim strPath As String

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   cp.Open "SELECT * FROM DMOVIN ORDER BY SOCIO,FECHA,APREPAC DESC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cl.MoveFirst

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

Private Sub CmdMovin_Click()

COLOCATITULOSENSOCIOS
'BorraCeldasenSOCIOS
If TxtclSocio = "50" Then
    Desglose
    Exit Sub
End If
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM DMOVIN ORDER BY SOCIO,FECHA,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    numes = 11
    Do Until cd.EOF = True
        If cd.Fields("SOCIO") = TxtclSocio Then
            'IntRespuesta = MsgBox("SOCIO=" & cd.Fields("CVEMOV") & cd.Fields("IMPORTE"), 0)

            If Month(cd.Fields("FECHA")) <> numes Then
                numes = Month(cd.Fields("FECHA"))
                RENGLON = RENGLON + 1
                Celda = RENGLON
                BorraCelda
            End If

            RENGLON = RENGLON + 1
            FGS.Col = 0
            FGS.Row = RENGLON
            FGS.Text = RENGLON
            FGS.Col = 1
            FGS.Text = cd.Fields("SOCIO")
            FGS.Col = 2
            FGS.Text = cd.Fields("FECHA")
            FGS.Col = 3
            FGS.Text = cd.Fields("CVEMOV")
            FGS.Col = 4
            FGS.Text = ""
            FGS.Col = 5
            FGS.Text = ""

            FGS.Col = 6
            FGS.Text = cd.Fields("DESCRIP")
            FGS.Col = 7
            If cd.Fields("REFERENC") > 0 Then
                FGS.Text = cd.Fields("REFERENC")
            End If
            FGS.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FGS.Text = cd.Fields("CTABCO")
            End If
            FGS.Col = 9
            FGS.Text = ""
            If cd.Fields("APREPAC") = "A" Then
                '      *Abonos
                FGS.Col = 9
                FGS.Text = Format(cd.Fields("IMPORTE"), "Currency")
                FGS.Col = 10
                FGS.Text = ""
                sdoActual = sdoActual + cd.Fields("IMPORTE")
            Else
                '      *Cargos
                FGS.Col = 9
                FGS.Text = ""
                FGS.Col = 10
                FGS.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual - cd.Fields("IMPORTE")
            End If
            FGS.Col = 11
            FGS.Text = Format(sdoActual, "Currency")
            
        End If

    cd.MoveNext
    
Loop
RenBorra = RENGLON
BorraCeldasenSOCIOS

cd.Close
TxtNumSocio.SetFocus
TxtNumSocio.SelStart = Val(TxtNumSocio.Text)
End Sub
Private Sub Desglose()
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'cd.Open "SELECT * FROM DMOVIN ORDER BY SOCIO,FECHA,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   cd.Open "SELECT * FROM DMOVIN ORDER BY SOCIO,APREPAC,CVEMOV,FECHA,REFERENC,IMPORTE", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    TotRetiros = 0

    Do Until cd.EOF = True
       If cd.Fields("SOCIO") = "50" Then
             'IntRespuesta = MsgBox("CVEMOV=" & cCveMov & "cd. " & cd.Fields("CVEMOV") & " " & cd.Fields("IMPORTE"), 0)

        If cd.Fields("CVEMOV") <> cCveMov Then

         If TotRetiros > 0 Then
            RENGLON = RENGLON + 1
            FGS.Col = 0
            FGS.Row = RENGLON
            FGS.Text = RENGLON
            FGS.Col = 1
            FGS.Text = ""
            FGS.Col = 2
            FGS.Text = ""
            FGS.Col = 3
            FGS.Text = ""
            FGS.Col = 4
            FGS.Text = ""
            FGS.Col = 5
            FGS.Text = ""

            FGS.Col = 6
            FGS.Text = "SUB-TOTAL"
            FGS.Col = 7
            FGS.Text = ""
            FGS.Col = 8
            FGS.Text = ""

            FGS.Col = 9
            FGS.Text = ""
            FGS.Col = 10
            FGS.Text = Format(TotRetiros, "Currency")

            RENGLON = RENGLON + 1
            FGS.Col = 0
            FGS.Row = RENGLON
            FGS.Text = RENGLON
            FGS.Col = 1
            FGS.Text = ""
            FGS.Col = 2
            FGS.Text = ""
            FGS.Col = 3
            FGS.Text = ""
            FGS.Col = 4
            FGS.Text = ""
            FGS.Col = 5
            FGS.Text = ""
         End If
         cCveMov = cd.Fields("CVEMOV")
         TotRetiros = 0

        End If

            RENGLON = RENGLON + 1
            FGS.Col = 0
            FGS.Row = RENGLON
            FGS.Text = RENGLON
            FGS.Col = 1
            FGS.Text = cd.Fields("SOCIO")
            FGS.Col = 2
            FGS.Text = cd.Fields("FECHA")
            FGS.Col = 3
            FGS.Text = cd.Fields("CVEMOV")
            FGS.Col = 4
            FGS.Text = ""
            FGS.Col = 5
            FGS.Text = ""

            FGS.Col = 6
            FGS.Text = cd.Fields("DESCRIP")
            FGS.Col = 7
            If cd.Fields("REFERENC") > 0 Then
                FGS.Text = cd.Fields("REFERENC")
            End If
            FGS.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FGS.Text = cd.Fields("CTABCO")
            End If
            FGS.Col = 9
            FGS.Text = ""
            If cd.Fields("APREPAC") = "A" Then
                '      *Abonos
                FGS.Col = 9
                FGS.Text = Format(cd.Fields("IMPORTE"), "Currency")
                FGS.Col = 10
                FGS.Text = ""
                sdoActual = sdoActual + cd.Fields("IMPORTE")
            Else
                '      *Cargos
                FGS.Col = 9
                FGS.Text = ""
                FGS.Col = 10
                FGS.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual - cd.Fields("IMPORTE")
                TotRetiros = TotRetiros + cd.Fields("IMPORTE")
            End If
            FGS.Col = 11
            FGS.Text = Format(sdoActual, "Currency")

        End If
                 cd.MoveNext

                 Loop
                 RENGLON = RENGLON + 1
            FGS.Col = 0
            FGS.Row = RENGLON
            FGS.Text = RENGLON
            FGS.Col = 1
            FGS.Text = ""
            FGS.Col = 2
            FGS.Text = ""
            FGS.Col = 3
            FGS.Text = ""
            FGS.Col = 4
            FGS.Text = ""
            FGS.Col = 5
            FGS.Text = ""

            FGS.Col = 6
            FGS.Text = "SUB-TOTAL"
            FGS.Col = 7
            FGS.Text = ""
            FGS.Col = 8
            FGS.Text = ""

            FGS.Col = 9
            FGS.Text = ""
            FGS.Col = 10
            FGS.Text = Format(TotRetiros, "Currency")
RenBorra = RENGLON
BorraCeldasenSOCIOS

cd.Close
TxtNumSocio.SetFocus
TxtNumSocio.SelStart = Val(TxtNumSocio.Text)
End Sub

Private Sub BorraCelda()
    RENGLON = Celda
    FGS.Row = RENGLON
    FGS.Text = RENGLON
    FGS.Col = 1
    FGS.Text = ""
    FGS.Col = 2
    FGS.Text = ""
    FGS.Col = 3
    FGS.Text = ""
    FGS.Col = 4
    FGS.Text = ""
    FGS.Col = 5
    FGS.Text = ""
    FGS.Col = 6
    FGS.Text = ""
    FGS.Col = 7
    FGS.Text = ""
    FGS.Col = 8
    FGS.Text = ""
    FGS.Col = 9
    FGS.Text = ""
    FGS.Col = 10
    FGS.Text = ""
    FGS.Col = 11
    FGS.Text = ""
End Sub




Private Sub CmdMovPres_Click()
Static lfrmCount As Long
    Dim frmD As FG
    lfrmCount = lfrmCount + 1
    Set frmD = New FG
    frmD.Caption = "FG"
    
    frmD.Show
End Sub

Private Sub CmdMovpr_Click()
COLOCATITULOSENSOCIOS
FGS.Col = 9
FGS.Text = "PAGOS"
FGS.Col = 10
FGS.Text = "PRESTAMOS"
'BorraCeldasenSOCIOS

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM DMOVPR ORDER BY SOCIO,FECHA,APREPAC DESC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

    RENGLON = 0
    cd.MoveFirst
    numes = 11
    Do Until cd.EOF = True
        If cd.Fields("SOCIO") = TxtclSocio Then
            If Month(cd.Fields("FECHA")) <> numes Then
                numes = Month(cd.Fields("FECHA"))
                RENGLON = RENGLON + 1
                Celda = RENGLON
                BorraCelda
            End If

            RENGLON = RENGLON + 1
            FGS.Col = 0
            FGS.Row = RENGLON
            FGS.Text = RENGLON
            FGS.Col = 1
            FGS.Text = cd.Fields("SOCIO")
            FGS.Col = 2
            FGS.Text = cd.Fields("FECHA")
            FGS.Col = 3
            FGS.Text = cd.Fields("CVEMOV")
            FGS.Col = 4
            FGS.Text = ""
            FGS.Col = 5
            FGS.Text = ""

            FGS.Col = 6
                FGS.Text = cd.Fields("DESCRIP")
            FGS.Col = 7
                If cd.Fields("REFERENC") > 0 Then
                    FGS.Text = cd.Fields("REFERENC")
                End If
                If cd.Fields("DESCRIP") = "CARGO POR INTERESES" Then
                    FGS.Text = Format(cd.Fields("TASA") / 100, "Percent")
                End If
            FGS.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FGS.Text = cd.Fields("CTABCO")
            End If
            FGS.Col = 9
            FGS.Text = ""
            FGS.Col = 10
            FGS.Text = ""
            If cd.Fields("CVEMOV") > "48" And cd.Fields("CVEMOV") < "60" Then
                  '*Abonos
                If cd.Fields("CVEMOV") = "50" And RENGLON < 2 Then
                    FGS.Col = 10
                    FGS.Text = Format(cd.Fields("IMPORTE"), "Currency")

                    sdoActual = sdoActual + cd.Fields("IMPORTE")
                Else
                    FGS.Col = 9
                    FGS.Text = Format(cd.Fields("IMPORTE"), "Currency")

                    sdoActual = sdoActual - cd.Fields("IMPORTE")
                End If
            Else
                '      *Cargos
                FGS.Col = 10
                FGS.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual + cd.Fields("IMPORTE")
            End If
            FGS.Col = 11
            FGS.Text = Format(sdoActual, "Currency")
        End If

    cd.MoveNext
Loop
RenBorra = RENGLON
BorraCeldasenSOCIOS

cd.Close
TxtNumSocio.SetFocus
TxtNumSocio.SelStart = Val(TxtNumSocio.Text)
End Sub

Private Sub CmdNvoSocio_Click()
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
        If cl.Fields("SOCIO") = TxtclSocio Then
           IntRespuesta = MsgBox(cl.Fields("SOCIO") & " YA EXISTE ESTE NUMERO", 0)
            Yaexiste = 1
            Exit Do
        Else
            cl.MoveNext
        End If
        If TxtclSocio = "" Then
            IntRespuesta = MsgBox("NO SE PUEDE CREAR UN SOCIO EN BLANCO. Digite un Número de Socio", 0)
            Yaexiste = 1
            Exit Do
        End If
        Loop
    If Yaexiste <> 1 Then
        With cl
        .AddNew
        cl.Fields("GRUPO") = TxtclGrupo
        cl.Fields("SOCIO") = TxtclSocio
        cl.Fields("TIPO") = TxtclTipo
        cl.Fields("PROMOTOR") = TxtclPromotor
        cl.Fields("CTASEGURO") = TxtSeguro
        cl.Fields("NOMBRE") = TxtclNombre
        cl.Fields("DIRECCION") = "D"
        cl.Fields("COLONIA") = "C"
        cl.Fields("CIUDAD") = "C"
        cl.Fields("ESTADO") = "E"
        cl.Fields("CP") = "0"
        cl.Fields("TELEFONO") = "5"
        cl.Fields("EMAIL") = "E"
        cl.Fields("CLABE") = TxtclCLABE
        cl.Fields("FECNAC") = txtclFecNac
        cl.Fields("FECAPER") = TxtclApertura
        cl.Fields("INV_INI") = "0"
        cl.Fields("APORTA") = "0"
        cl.Fields("RETIROS") = "0"
        cl.Fields("SALDO") = "0"
        cl.Fields("INTGANADO") = "0"
        cl.Fields("COMISION") = "0"
        cl.Fields("PRES_INI") = "0"
        cl.Fields("PRESTAMOS") = "0"
        cl.Fields("PAGOS") = "0"
        cl.Fields("INTPAGADO") = "0"
        cl.Fields("SALDOPRES") = "0"
        cl.Fields("FECPRES") = TxtFPrestamo
        cl.Fields("FECVENC") = TxtFVencimiento
        cl.Fields("PAGOMIN") = 0
        cl.Fields("TASA_INV") = 0
        cl.Fields("TASAPRES") = TxtTasaPres
        cl.Fields("FECORTE") = Date
        
        cl.Update
        End With
   End If
End Sub




Private Sub FGS_Click()
    FGS.Col = 1
    PrvSocio = FGS.Text
    FGS.Col = 2
    PrvFecha = FGS.Text
    FGS.Col = 3
    PrvCveMov = FGS.Text
    FGS.Col = 5
    PrvAPrePac = FGS.Text
    FGS.Col = 7
    PrvReferenc = FGS.Text
    FGS.Col = 9
    If FGS.Text <> "" Then
        prvImporte = FGS.Text
    Else
        FGS.Col = 10
        prvImporte = FGS.Text
    End If
    TxtclSocio = PrvSocio
    'IntRespuesta = MsgBox(prvImporte, 0)
    TxtNomSocio = ""
End Sub

Private Sub Form_Load()
'   IntRespuesta = MsgBox("Mi Primera Carpeta=" & frmMiPrimera.LblCarpeta, 0)
    Carpeta = frmMiPrimera.LblCarpeta
    PrvSocio = frmMiPrimera.LblSocio
    'TxtNumSocio = "01"
    TxtclSocio = frmMiPrimera.LblSocio
    'SendKeys "(Enter)"
    'ImgCaptura.Picture = LoadPicture("C:\" & Carpeta & "\Happy.bmp")

    'IntRespuesta = MsgBox("Carpeta=" & PrvSocio, 0)

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb" '''
 
 End Sub

Private Sub TxtclGrupo_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtclGrupo.Tag = MODE_OVERTYPE And TxtclGrupo.SelLength = 0 Then
        TxtclGrupo.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       SendKeys "{tab}"
    End If
    If flgsocio <> 1 Then
        If KeyAscii <> 13 Then
            flgsocio = 1
            TxtclGrupo = ""
        End If
    End If
End Sub

Private Sub TxtclNombre_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 97) And (KeyAscii <= 122) Then
        KeyAscii = KeyAscii - 32
    End If
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtclNombre.Tag = MODE_OVERTYPE And TxtclNombre.SelLength = 0 Then
        TxtclNombre.SelLength = 1
    End If
    
    If flgsocio = 0 Then
        TxtclNombre = ""
    End If
    
    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       SendKeys "{tab}"
    End If
    If flgsocio <> 1 Then
        If KeyAscii <> 13 Then
            flgsocio = 1
            TxtclNombre = ""
        End If
    End If
End Sub

Private Sub TxtclPromotor_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtclPromotor.Tag = MODE_OVERTYPE And TxtclPromotor.SelLength = 0 Then
        TxtclPromotor.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       SendKeys "{tab}"
    End If
    If flgsocio <> 1 Then
        If KeyAscii <> 13 Then
            flgsocio = 1
            TxtclPromotor = ""
        End If
    End If
End Sub

Private Sub TxtclSocio_Change()

          TxtPagoMin = 0

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'ESTA LINEA SOBRA porque ya has especificado el origen en la SELECT:
    'cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"

   'Esto tiene que funcionar
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   'cl.Open "SELECT * FROM CLIENTES ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   cl.MoveFirst
    'TxtclSocio.Text = pubSocio
    'IntRespuesta = MsgBox(PubSocio, 0)

    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = TxtclSocio Then
          PrvSocio = TxtclSocio
          PubSocio = TxtclSocio
          TxtclSocio.Text = PrvSocio
          TxtclTipo = cl.Fields("TIPO")
          frmMiPrimera.LblSocio = PrvSocio
          PubNombre = cl.Fields("NOMBRE")
          TxtclGrupo = cl.Fields("GRUPO")
          PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\LOBAHNOS.jpg")

          'PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Fedamac.jpg")
          'MsgBox (Fotosocio)
          If TxtclTipo < "3" Then
              Fotosocio = "Foto" & TxtclSocio & ".jpg"
          Else
              Fotosocio = "Foto" & TxtclGrupo & ".jpg"
          End If
        
          PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\" & Fotosocio)

        
          TxtclTipo = cl.Fields("TIPO")
          TxtclPromotor = cl.Fields("PROMOTOR")
          TxtclNombre = cl.Fields("NOMBRE")
          TxtclDireccion = cl.Fields("DIRECCION")
          TxtclColonia = cl.Fields("COLONIA")
          TxtclCiudad = cl.Fields("CIUDAD")
          TxtclEstado = cl.Fields("ESTADO")
          TxtclCP = cl.Fields("CP")
          TxtclTelefono = cl.Fields("TELEFONO")
          TxtcleMail = cl.Fields("EMAIL")
          txtclFecNac = cl.Fields("FECNAC")
          TxtclApertura = cl.Fields("FECAPER")
          TxtSaldoInicial = Format(cl.Fields("INV_INI"), "Currency")
          TxtAportaciones = Format(cl.Fields("APORTA"), "Currency")
          TxtRetiros = Format(cl.Fields("RETIROS"), "Currency")
          TxtSaldoActual = Format(cl.Fields("SALDO"), "Currency")
          TxtIntsDevengados = Format(cl.Fields("INTGANADO"), "Currency")
          TxtComisiones = Format(cl.Fields("COMISION"), "Currency")
          TxtPrestamoInicial = Format(cl.Fields("PRES_INI"), "Currency")
          TxtPrestamos = Format(cl.Fields("PRESTAMOS"), "Currency")
          TxtPagos = Format(cl.Fields("PAGOS"), "Currency")
          TxtIntsPagados = Format(cl.Fields("INTPAGADO"), "Currency")
          TxtPrestamoActual = Format(cl.Fields("SALDOPRES"), "Currency")
          varPagoTotal = Format(cl.Fields("SALDOPRES") + (cl.Fields("SALDOPRES") * cl.Fields("TASAPRES") / 100), "Currency")
          TxtPagoTotal = varPagoTotal
          TxtFechaCorte = cl.Fields("FECORTE")
          If Month(cl.Fields("FECORTE")) > 10 Then
            varNumMes = Month(cl.Fields("FECORTE")) - 10
          Else
            varNumMes = Month(cl.Fields("FECORTE")) + 2
          End If
          TxtSeguro = cl.Fields("CTASEGURO")
          If cl.Fields("ULTPAGO") <> "" Then
             TxtUltPago = cl.Fields("ULTPAGO")
          End If
          If cl.Fields("CLABE") <> "" Then
             CLABE = cl.Fields("CLABE")
          End If
          'IntRespuesta = MsgBox(CLABE, 0)

          TxtclCLABE = CLABE
          
          
              
         If cl.Fields("APORTA") > 0 Then
            TxtPromAportacion = Format(cl.Fields("APORTA") / varNumMes, "Currency")
         End If
            TxtPromInversion = Format(cl.Fields("PROM_INV"), "Currency")
          
         If cl.Fields("PROM_INV") > 0 Then
            TxtTasaInversion = Format(cl.Fields("INTGANADO") / cl.Fields("PROM_INV"), "Percent")
         Else
            TxtTasaInversion = ""
         End If
       
          TxtInvMoroleon = Format(cl.Fields("MORINI"), "Currency")

         'If cl.Fields("SALDOPRES") > 0 Then
            TxtFVencimiento = cl.Fields("FECVENC")
            TxtFPrestamo = cl.Fields("FECPRES")
            TxtTasaPres = cl.Fields("TASAPRES")
            TxtPagoMin = Format(cl.Fields("PAGOMIN"), "Currency")
         'Else
         '   TxtFVencimiento = Date
         '   TxtFPrestamo = Date
         '   TxtTasaPrestamo = 0
         '   TxtPagoMinimo = 0
         'End If
         TxtSaldoCaja = Format(TxtSaldoActual - TxtPrestamoActual, "Currency")
         'IntRespuesta = MsgBox(cl.Fields("FECNAC"), 0)
          

          Exit Do
       Else
          TxtclNombre = "No existe nombre de este Socio"
          TxtclGrupo = ""
          TxtclTipo = ""
          TxtclPromotor = ""
          TxtclDireccion = "D"
          TxtclColonia = "C"
          TxtclCiudad = "C"
          TxtclEdo = "E"
          TxtclCP = "0"
          TxtclTelef = "5"
          TxtcleMail = "E"
          TxtclApertura = Date
          TxtSeguro = ""
          txtclFecNac = Date
          TxtclApertura = Date
          TxtSaldoInicial = 0
          TxtAportaciones = 0
          TxtRetiros = 0
          TxtSaldoActual = 0
          TxtIntsDevengados = 0
          TxtComisiones = 0
          TxtPrestamoInicial = 0
          TxtPrestamos = 0
          TxtPagos = 0
          TxtIntsPagados = 0
          TxtPrestamoActual = 0
          varPagoTotal = 0
          TxtPagoTotal = 0
          TxtFechaCorte = Date
          varNumMes = 0
          TxtclCLABE = CLABE
          TxtPromAportacion = 0
          TxtPromInversion = 0
          TxtTasaInversion = 0
          TxtInvMoroleon = 0
          TxtFVencimiento = Date
          TxtFPrestamo = Date
          TxtTasaPres = 0
          TxtPagoMinimo = 0
          TxtSaldoCaja = 0
          'IntRespuesta = MsgBox(cl.Fields("FECNAC"), 0
       End If
       cl.MoveNext
    Loop
'cl.AddNew

cl.Close

End Sub
Sub BorraCeldasenFG()
RENGLON = RenBorra
    Do Until RENGLON = 199
       RENGLON = RENGLON + 1
       FGS.Col = 0
       FGS.Row = RENGLON
       FGS.Text = ""
       FGS.Col = 1
       FGS.Text = ""
       FGS.Col = 2
       FGS.Text = ""
       FGS.Col = 3
       FGS.Text = ""
       FGS.Col = 4
       FGS.Text = ""
       FGS.Col = 5
       FGS.Text = ""
       FGS.Col = 6
       FGS.Text = ""
       FGS.Col = 7
       FGS.Text = ""

    Loop

End Sub


Private Sub TxtclTipo_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtclTipo.Tag = MODE_OVERTYPE And TxtclTipo.SelLength = 0 Then
        TxtclTipo.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       SendKeys "{tab}"
    End If
    If flgsocio <> 1 Then
        If KeyAscii <> 13 Then
            flgsocio = 1
            TxtclTipo = ""
        End If
    End If
End Sub

Private Sub TxtNomSocio_Change()

'COLOCATITULOSENMS
'Sub COLOCATITULOSENMS()
FGS.Row = 0
FGS.Col = 0
FGS.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
FGS.Col = 1
FGS.Text = "SOCIO"
FGS.Col = 2
FGS.Text = "GRUPO"
FGS.Col = 3
FGS.Text = "NOMBRE"
FGS.Col = 4
FGS.Text = "SALDO"
FGS.Col = 5
FGS.Text = "PRESTAMO"
FGS.Col = 6
FGS.Text = "INTERESES"
FGS.Col = 7
FGS.Text = "COMISION"

FGS.ColWidth(3) = 3000    'AJUSTO EL ANCHO DE LA COLUMNA2
FGS.ColWidth(4) = 1000    'AJUSTO EL ANCHO DE LA COLUMNA2
FGS.ColWidth(5) = 1000    'AJUSTO EL ANCHO DE LA COLUMNA2

RENGLON = 0
'FGS.Row = 1

'End Sub
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'ESTA LINEA SOBRA porque ya has especificado el origen en la SELECT:
    'cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"

   'Esto tiene que funcionar
   cl.Open "SELECT * FROM SOCIOS ORDER BY NOMBRE", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
   cl.MoveFirst
    LongNom = 1
    LongNom = Len(TxtNomSocio)
    Do Until cl.EOF = True
        varnombre = Left(cl.Fields("NOMBRE"), LongNom)
        UNOMBRE = TxtNomSocio
        UNOMBRE = UCase(UNOMBRE)

        If varnombre = UNOMBRE Then
            RENGLON = RENGLON + 1
            FGS.Col = 0
            FGS.Row = RENGLON
            FGS.Text = RENGLON
            FGS.Col = 1
            FGS.Text = cl.Fields("SOCIO")
            FGS.Col = 2
            FGS.Text = cl.Fields("GRUPO")
            FGS.Col = 3
            FGS.Text = cl.Fields("NOMBRE")
            FGS.Col = 4
            FGS.Text = Format(cl.Fields("SALDO"), "Currency")
            FGS.Col = 5
            FGS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
        End If
        cl.MoveNext
    Loop

cl.Close
'ValorFlexGrid
RenBorra = RENGLON
BorraCeldasenFG

End Sub
Private Sub TxtNumSocio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & TxtNumSocio, 0)
       TxtclSocio = TxtNumSocio
       FlexGridSocio
       'SendKeys "(tab)"
       'Exit Sub
    End If

End Sub
Private Sub FlexGridSocio()
FGS.Row = 0
FGS.Col = 0
FGS.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
FGS.Col = 1
FGS.Text = "SOCIO"
FGS.Col = 2
FGS.Text = "GRUPO"
FGS.Col = 3
FGS.Text = "NOMBRE"
FGS.Col = 4
FGS.Text = "SALDO"
FGS.Col = 5
FGS.Text = "PRESTAMO"
FGS.Col = 6
FGS.Text = "INTERESES"
FGS.Col = 7
FGS.Text = "COMISION"

FGS.ColWidth(3) = 3000    'AJUSTO EL ANCHO DE LA COLUMNA2
FGS.ColWidth(4) = 1000    'AJUSTO EL ANCHO DE LA COLUMNA2
FGS.ColWidth(5) = 1000    'AJUSTO EL ANCHO DE LA COLUMNA2
    'Do Until RENGLON = 199
    '   RENGLON = RENGLON + 1
    '   FGS.Col = 0
    '   FGS.Row = RENGLON
    '   FGS.Text = ""
     '  FGS.Col = 1
     '  FGS.Text = ""
     '  FGS.Col = 2
     '  FGS.Text = ""
     '  FGS.Col = 3
     '  FGS.Text = ""
     '  FGS.Col = 4
     '  FGS.Text = ""
     '  FGS.Col = 5
     '  FGS.Text = ""
     '  FGS.Col = 6
     '  FGS.Text = ""
     '  FGS.Col = 7
     '  FGS.Text = ""

    'Loop
RENGLON = 0
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'ESTA LINEA SOBRA porque ya has especificado el origen en la SELECT:
    'cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"

   'Esto tiene que funcionar
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
   cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") >= TxtNumSocio Then

            RENGLON = RENGLON + 1
            FGS.Col = 0
            FGS.Row = RENGLON
            FGS.Text = RENGLON
            FGS.Col = 1
            FGS.Text = cl.Fields("SOCIO")
            FGS.Col = 2
            FGS.Text = cl.Fields("GRUPO")
            FGS.Col = 3
            FGS.Text = cl.Fields("NOMBRE")
            FGS.Col = 4
            FGS.Text = Format(cl.Fields("SALDO"), "Currency")
            FGS.Col = 5
            FGS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
            FGS.Col = 6
            FGS.Text = ""
            FGS.Col = 7
            FGS.Text = ""
            FGS.Col = 8
            FGS.Text = ""
            If RENGLON > 15 Then
                TxtNumSocio = ""
                Exit Sub
            End If
      End If
    cl.MoveNext
Loop

TxtNumSocio = ""
RenBorra = RENGLON
BorraCeldasenSOCIOS
cl.Close

End Sub


Private Sub TxtSeguro_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtSeguro.Tag = MODE_OVERTYPE And TxtSeguro.SelLength = 0 Then
        TxtSeguro.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       SendKeys "{tab}"
       flgsocio = 0
    End If
    If flgsocio <> 1 Then
        If KeyAscii <> 13 Then
            flgsocio = 1
            TxtSeguro = ""
        End If
    End If
End Sub

VERSION 5.00
Begin VB.Form FrmCapitalizar 
   BackColor       =   &H8000000E&
   Caption         =   "Capitalizar"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtSaldoCaja 
      Height          =   285
      Left            =   4200
      TabIndex        =   27
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox TxtFactor 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000000000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   26
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox TxtIntGanado 
      Height          =   285
      Left            =   4200
      TabIndex        =   25
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox TxtTotComision 
      Height          =   285
      Left            =   4200
      TabIndex        =   23
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox TxtReserva 
      Height          =   285
      Left            =   4200
      TabIndex        =   21
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox TxtTotalIntPagados 
      Height          =   285
      Left            =   4200
      TabIndex        =   19
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox TxtTotalPromedios 
      Height          =   285
      Left            =   4200
      TabIndex        =   18
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox TxtInicioEjercicio 
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Text            =   "01/11/2011"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton CmdCapitalizar 
      Caption         =   "Proceso de Capitalización"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox TxtFinEjercicio 
      Height          =   285
      Left            =   3480
      TabIndex        =   14
      Text            =   "31/10/2012"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox TxtFecorte 
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Text            =   "30/11/2010"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label LblEspere 
      Height          =   375
      Left            =   4080
      TabIndex        =   28
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label15 
      Caption         =   "Intereses Ganados"
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label14 
      Caption         =   "Total de Comisiones"
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "Reserva de Capital"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Inicio de Ejercicio"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Fin de Ejercicio"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label LblMesActual 
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label LblMesAnterior 
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Saldo Actual"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Saldo Anterior"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Fecha de Corte"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Proceso de Capitalización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   6
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Saldo en Caja"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Factor de Prorrateo"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Total Promedio de Inversión"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Total Intereses Pagados"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label LblProceso 
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX°°°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
      Width           =   8775
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "FrmCapitalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Carpeta As String
Private x, sa_cero As Integer
Private PrvFecorte, s_fecha As Date
Private PrvNumovs, s_prestamos, t_intpagado, t_pagos, t_prestamos, x_intpagados, IntsInv As Single
Private IntRespuesta, s_promotor, s_tercero As String
Private PrvMesCorte, MesCorte, r_pagos As String
Private s_pagos, s_prestamo, s_cargoint, c_abonos, c_cargos, acum_int, s_prominv, ac_prominv As Single
Private c_intbanco, c_isrbanco, c_invfin, c_prestamo, c_fondo, s_aporta, aporta50 As Single
Private s_retiros, s_presta, s_pagopres, s_intpagado, s_saldopres, s_presfina, v_mes As Single
Private s_socio, s_grupo, cs_socio As String
Private s_invini, s_tasa, s_corte, cs_importe, cs_prestamo, s_comision, TotComision As Single
Private s_prestini, s_invfin, s_presfin, s_intinver, s_Intpres As Single
Private s_apfundac, s_saldisp, s_salcaja, s_factor, s_intganado As Single
Private s_fecpres As String
Private diad As Date
Private t_prominv, s_saldoActual, s_saldomesa As Single
Private s_tipo, no_pago As String
Private s_aprepac As String
Private s_cvemov As String
Private s_importe, flg_next, flg_61, pago_anticipado As Single
Private PrvNombre As String
Private PrvPagoMin, PrvTasaPres As Single
Private s_numes, mes_pre As Single
Private s_descrip As String
Private s_referenc, sema_id, cl_sema_id As String
Private s_delete As Single


Option Explicit

Const MODE_OVERTYPE = "overtype"
Const MODE_INSERT = "insert"



Private Sub Form_Load()
    flg_next = 0
    flg_61 = 0
    'IntRespuesta = MsgBox("Mi Primera Carpeta=" & frmMiPrimera.LblCarpeta, 0)
    LblEmpresa = frmMiPrimera.LblEmpresa
    Carpeta = frmMiPrimera.LblCarpeta
    'IntRespuesta = MsgBox("Carpeta=" & Carpeta, 0)

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb" '''
   
    TxtFecorte.Tag = MODE_OVERTYPE
    LblProceso.Tag = MODE_OVERTYPE
    
'Busca la fecha más reciente de los movimientos

   Dim cs As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim cb As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cs.Open "SELECT * FROM SICMOV ORDER BY FECHA", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   cb.Open "SELECT * FROM MPBXCTRL", cn
   
    cs.MoveFirst
    
    Do Until cs.EOF = True
        PrvNumovs = PrvNumovs + 1
        PrvFecorte = cs.Fields("FECHA")
        LblProceso.Caption = PrvNumovs & " Registros. Ultimo Registro = $" & cs.Fields("IMPORTE")
        cs.MoveNext
        Loop
        'IntRespuesta = MsgBox("PrvNumovs = " & PrvNumovs, 0)
TxtFecorte = PrvFecorte
MesCorte = Month(PrvFecorte)
PrvMesCorte = Month(PrvFecorte) + 15
If MesCorte > 10 Then
    MesCorte = MesCorte - 10
Else
    MesCorte = MesCorte + 2
End If
    IntRespuesta = MsgBox("MesCorte = " & MesCorte, 0)
    PrvMesCorte = MesCorte + 16

cs.Close
s_socio = "99"
BUSCA_SOCIO
       IntRespuesta = MsgBox("PrvNumovs = " & PrvNumovs & PrvFecorte, 0)

cb.MoveFirst
sema_id = "SA " & Str(Month(TxtFecorte) - 1)
Do Until cb.EOF = True

    If cb.Fields("SEMA_ID") = sema_id Then
        LblMesAnterior.Caption = Format(cb.Fields("DATA_N"), "Currency")

        'cb.MovePrevious
        'IntRespuesta = MsgBox("PrvMesCorte= " & PrvMesCorte, 0)
        cb.MoveNext
        LblMesActual.Caption = Format(cb.Fields("DATA_N"), "Currency")
        LblMesActual.Caption = Format(s_saldoActual, "Currency")

        Exit Do
    End If
    cb.MoveNext
    Loop

cb.Close
End Sub


Private Sub TxtFecorte_Change()
    'If TxtFecorte.Tag = MODE_OVERTYPE And TxtFecorte.SelLength = 0 Then
        TxtFecorte.SelLength = 1
    'End If

End Sub

'*Programa: SICAPITA.PRG
'*Fecha: 01-dic-2004-14/08/2009.-capitaliza inversi¢n 07/11/2011
'*Autor: Luis L¢pez Baeza
'*M¢dulo de C lculo de Intereses sobre Pr‚stamos del Fondo FEDAMAC
'************************************************************************
Private Sub CmdCapitalizar_Click()
    
    Dim stipo As String
    diad = TxtFinEjercicio
    IntRespuesta = MsgBox("Proceso de Capitalización. Deseas Continuar...?", 1)
    If (IntRespuesta = 1) Then
        IntRespuesta = MsgBox("Continúa", 0)
    Else
        Exit Sub
    End If
'Borra_DETMOV'ABRE ("SICMOV")

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   
   'Dim cd As New ADODB.Recordset 'Creamos el Objeto Recordset.DETMOV
   Dim cp As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVPR
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS
   Dim cs As New ADODB.Recordset 'Creamos el objeto Recordser.SICMOV
   Dim cv As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVIN
   
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."
   cs.Open "SELECT * FROM SICMOV ORDER BY APREPAC,SOCIO,FECHA,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

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
BORRA_DMOVIN
'BORRA_DETMOV
BORRA_DMOVPR
BORRA_SOCIOS

'IntRespuesta = MsgBox("PREPARA ARCHIVOS DMOVIN Y DMOVPR", 0)



'*Inicializa Variables
s_pagos = 0
s_prestamo = 0
s_cargoint = 0
c_abonos = 0
c_cargos = 0
acum_int = 0
s_prominv = 0
ac_prominv = 0
c_intbanco = 0
c_isrbanco = 0
c_invfin = 0
c_prestamo = 0
c_fondo = 0
s_aporta = 0
aporta50 = 0
s_retiros = 0
s_presta = 0
s_pagopres = 0
s_intpagado = 0
s_saldopres = 0
s_presfina = 0
v_mes = 1 '


     LblProceso.Caption = "Prepara Socios y Saldo Inicial del Ejercicio"
    
    


LblProceso.Caption = PrvNumovs & " Calcula Promedios de Inversión" '


    CREA_PROMEDIOS

TxtTotalPromedios = Format(ac_prominv, "Currency")
LblProceso.Caption = "INICIA CALCULO DE INTERESES SOBRE PRESTAMOS"
    LblEspere.Caption = "       Espere un momento por favor..."

'LblEspere.Caption = ""
'LblProceso.Caption = "                                 Espere un momento por favor...."

IntRespuesta = MsgBox("EMPIEZA CALCULO DE INTERESES SOBRE PRESTAMOS", 0)
    'LblEspere.Caption = "Espere un momento por favor..."

cs.Close
cs.Open "SELECT * FROM SICMOV ORDER BY SOCIO,NUMES,FECHA,APREPAC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.


  
   cs.MoveFirst
    
   s_numes = 0
   s_socio = "01"
   no_pago = "0"
   s_tasa = 1
   r_pagos = "0"
   
   s_saldopres = 0
   pago_anticipado = 0
    Do Until cs.EOF = True
    x = x + 3
    Me.CurrentX = x
    Me.CurrentY = 6
    Me.Print "***"
          If cs.Fields("APREPAC") = "C" Or cs.Fields("APREPAC") = "P" Then

             If cs.Fields("SOCIO") <> cs_socio Then
                sa_cero = 0
                
                s_socio = cs_socio

                'If s_saldopres > 0 Or s_saldopres < 0 Then
                   s_saldomesa = s_saldopres + s_pagos
                   If flg_61 = 1 Then
                       s_saldomesa = s_saldopres - cs_prestamo + s_pagos
                   End If
                   s_corte = 1
                   CORTE_MES
                   's_saldopres = s_saldopres - s_pagos
                   '+ cs_prestamo
                   s_pagos = 0
                'End If
                Do While s_numes < MesCorte
                    s_corte = 1.2
                    s_numes = s_numes + 1
                    s_saldomesa = s_saldopres
                    '- cs_prestamo


                    If s_saldomesa > 0 Then
                        If flg_61 = 0 Then
                            flg_61 = 0
                            s_saldomesa = s_saldopres - cs_prestamo
                        End If
                        
                        '+ s_pagos
                        CORTE_MES
                    End If
                    s_pagos = 0
                Loop
                s_socio = cs_socio
                s_saldopres = 0
                cs_prestamo = 0
                If cs.Fields("APREPAC") = "C" And (cs.Fields("CVEMOV") = "61" Or cs.Fields("CVEMOV") = "62") Or cs.Fields("CVEMOV") = "63" Then
                    If sa_cero = 0 Then
                        sa_cero = 1
                        s_numes = cs.Fields("NUMES")
                    End If
                    cs_importe = cs.Fields("IMPORTE")
                    cs_socio = cs.Fields("SOCIO")
                    If cs.Fields("NUMES") <= s_numes Then
                        s_numes = cs.Fields("NUMES")
                    End If
                    If s_numes = 1 Then
                        s_numes = cs.Fields("NUMES")
                        mes_pre = cs.Fields("NUMES")
                    End If
                    If flg_next = 0 Then
                        cs_prestamo = cs_prestamo + cs_importe
                        s_saldopres = s_saldopres + cs_importe
                    End If
                    flg_61 = 1

                End If
             Else
                If s_numes <> cs.Fields("NUMES") Then
'If cs_socio = "15" Then
'   IntRespuesta = MsgBox("Importe=" & cs.Fields("IMPORTE") & "S_SALDOPRES=" & s_saldopres & " s_pagos=" & s_pagos & " cs_prestamo=" & cs_prestamo & "s_sadomesa=" & s_saldomesa & " s_numes=" & s_numes & " flg_next=" & flg_next & flg_61, 0)
'End If

                    s_saldomesa = s_saldopres + s_pagos
                    '- cs_prestamo
                    If flg_61 = 1 Then
                        s_saldomesa = s_saldopres + s_pagos - cs_prestamo
                    End If
                    If flg_next = 1 Then
                       cs_importe = s_pagos
                       s_saldomesa = s_saldopres - cs_prestamo
                       'cs_prestamo = 0
                       s_pagos = 0
                    End If

                    s_corte = 2
                    If s_saldomesa > 0 Then
                        CORTE_MES
                    End If
                    If flg_61 = 1 Then
                        cs_prestamo = 0
                        flg_61 = 0
                    End If
                    'If flg_next = 1 Then
                    '   s_pagos = cs_importe
                    'End If
                    's_saldomesa = s_saldopres - s_pagos + cs_prestamo
                    s_numes = s_numes + 1
                    s_socio = cs_socio
                    s_pagos = 0
                    If flg_next = 1 Then
                        flg_next = 0
                    End If
                    If s_numes < cs.Fields("NUMES") Then
                        If cs.Fields("APREPAC") = "P" Then
                           flg_next = 1
                        End If
                    End If
                    If s_saldopres <> cs_prestamo Then
                        cs_prestamo = 0
                    End If
                End If

                If cs.Fields("APREPAC") = "P" Then
                    If cs.Fields("CVEMOV") = "54" Then
                        pago_anticipado = 1
                    End If
                    If cs.Fields("SOCIO") = cs_socio Then
                        cs_importe = cs.Fields("IMPORTE")
                        If flg_next = 0 Then
                            s_saldopres = s_saldopres - cs_importe
                            s_pagos = s_pagos + cs_importe
                        End If
                    End If
                End If

                If cs.Fields("APREPAC") = "C" And (cs.Fields("CVEMOV") = "61" Or cs.Fields("CVEMOV") = "62") Or cs.Fields("CVEMOV") = "63" Then
                    If sa_cero = 0 Then
                        sa_cero = 1
                        s_numes = cs.Fields("NUMES")
                        s_saldomesa = 0
                    End If
  
                    cs_importe = cs.Fields("IMPORTE")
                    If cs.Fields("NUMES") <= s_numes Then
                        s_numes = cs.Fields("NUMES")
                    End If
                    If s_numes = 1 Then
                        s_numes = cs.Fields("NUMES")
                        mes_pre = cs.Fields("NUMES")
                    End If
                    If flg_next = 0 Then
                        cs_prestamo = cs_prestamo + cs_importe
                        s_saldopres = s_saldopres + cs_importe
                    End If
                    
'If cs.Fields("SOCIO") = "11" Then
'   IntRespuesta = MsgBox("APREPAC C" & cs.Fields("IMPORTE") & "S_SALDOPRES=" & s_saldopres & " s_pagos=" & s_pagos & " cs_prestamo=" & cs_prestamo & "s_sadomesa=" & s_saldomesa & " flg_next=" & flg_next & flg_61, 0)
'End If
                    flg_61 = 1

                End If
                
                
             End If
            
             If cs.Fields("APREPAC") = "C" And cs.Fields("DESCRIP") = "SALDO ANTERIOR" Then
                If cs.Fields("IMPORTE") = 0 Then
                    sa_cero = 0
                Else
                    sa_cero = 1
                End If
                s_saldopres = 0
                cs_importe = 0
                cs_prestamo = 0
                s_saldopres = cs.Fields("IMPORTE")
                cs_socio = cs.Fields("SOCIO")
                s_socio = cs_socio
                s_numes = cs.Fields("NUMES")
                s_pagos = 0
                If Not cs.EOF Then
                    'cs.MoveNext
                Else
                    Exit Do
                End If
                
             End If
             If cs.EOF Then
                Exit Do
             End If
             
            
          Else
             'cs.MoveNext
          End If
          If flg_next = 0 Then
            cs.MoveNext
          End If
    Loop
    x_intpagados = x_intpagados + IntsInv
    TxtTotalIntPagados = Format(x_intpagados, "currency")
    TxtTotComision = Format(TotComision, "currency")
    TxtReserva = Format((x_intpagados - TotComision) * 0.15, "currency")
    TxtIntGanado = Format(x_intpagados - TotComision - TxtReserva, "currency")
    s_factor = TxtIntGanado / TxtTotalPromedios
    TxtFactor = s_factor
    TxtSaldoCaja = LblMesActual.Caption
    'IntRespuesta = MsgBox(TxtIntGanado & TxtTotalPromedios & TxtFactor, 0)
    LblProceso.Caption = "PRORRATEO DE INTERESES DEVENTADOS"

    'LblProceso.Caption = "TERMINO CALCULO DE INTERESES SOBRE PRESTAMOS"
LblEspere.Caption = ""
    IntRespuesta = MsgBox("PRORRATEO DE INTERESES DEVENGADOS...OK", 0)
        LblEspere.Caption = "Espere un momento por favor..."

    
    PRORRATEO_INTERESES_SOCIO
        x = x + 3
    Me.CurrentX = x
    Me.CurrentY = 6
    Me.Print
    Me.Print
    Me.Print "TERMINO PROCESO DE CAPITALIZACIION"
    IntRespuesta = MsgBox("TERMINÓ PRORRATEO DE INTERESES DEVENGADOS...OK", 0)
    LblProceso.Caption = "TERMINO PROCESO DE CAPITALIZACION...OK"
LblEspere.Caption = ""
    IntRespuesta = MsgBox("TERMINÓ PROCESO DE CAPITALIZACION...OK", 0)
    ACTUALIZA_MPBXCTRL
 
End Sub
Private Sub ACTUALIZA_MPBXCTRL()
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   Dim strPath As String

    strPath = "C:\" & Carpeta & "\SISFED.mdb"

    Set cn = New ADODB.Connection

    With cn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM MPBXCTRL"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    cl.MoveFirst
    sema_id = "SA " & Str(Month(TxtFecorte))

    Do Until cl.EOF = True
    x = x + 3
    Me.CurrentX = x
    Me.CurrentY = 6
    Me.Print ">>>"
        cl_sema_id = cl.Fields("SEMA_ID")
        If cl_sema_id = sema_id Then
            cl.Fields("DATA_N") = LblMesActual
            cl.Fields("DATA_D") = TxtFecorte
            cl.Update
            Exit Sub
        End If
        
        cl.MoveNext
    Loop
End Sub

Private Sub PRORRATEO_INTERESES_SOCIO()

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   Dim strPath As String

    strPath = "C:\" & Carpeta & "\SISFED.mdb"

    Set cn = New ADODB.Connection

    With cn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM SOCIOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    cl.MoveFirst
    Do Until cl.EOF = True
    x = x + 3
    Me.CurrentX = x
    Me.CurrentY = 6
    Me.Print ">>>"
        If cl.Fields("SOCIO") = "50" Then
            cl.Fields("COMISION") = TxtReserva
        End If
        s_socio = cl.Fields("SOCIO")
        s_grupo = cl.Fields("GRUPO")
        s_tipo = cl.Fields("TIPO")
        s_promotor = cl.Fields("PROMOTOR")
        s_intganado = cl.Fields("PROM_INV") * s_factor
        If s_tipo = "3" Then
            s_tasa = 3
            s_comision = s_intganado * 0.1
            s_intganado = s_intganado - s_comision
        End If
        cl.Fields("INTGANADO") = s_intganado
        cl.Update
        If cl.Fields("INTGANADO") > 0 Then
            s_prominv = cl.Fields("PROM_INV")
            s_intganado = cl.Fields("INTGANADO")
            CREA_ABONO_POR_INTGANADO
        End If
        cl.MoveNext
    Loop
End Sub
Private Sub CREA_ABONO_POR_INTGANADO()
    
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cv As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVIN

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
 Set cv = New ADODB.Recordset

    With cv
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVIN"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

    With cv
        .AddNew
        cv.Fields("NUMREG") = cv.Fields("Id")
        cv.Fields("TIPO") = "T"
        cv.Fields("APREPAC") = "A"
        cv.Fields("CVEMOV") = "99"
        cv.Fields("IMPORTE") = s_intganado
        cv.Fields("SOCIO") = s_socio
        cv.Fields("GRUPO") = s_grupo
        cv.Fields("FECHA") = PrvFecorte
        cv.Fields("NUMES") = s_numes
        cv.Fields("DESCRIP") = "INTERESES DEVENGADOS"
        cv.Fields("TASA") = s_intganado / s_prominv
        cv.Update
        If s_tipo = "3" Then
            s_tasa = 3
            'CREA_ABONO_POR_COMISION
            .AddNew
            cv.Fields("NUMREG") = cv.Fields("Id")
            cv.Fields("TIPO") = "T"
            cv.Fields("APREPAC") = "A"
            cv.Fields("CVEMOV") = "99"
            cv.Fields("IMPORTE") = s_comision
            cv.Fields("SOCIO") = s_promotor
            cv.Fields("GRUPO") = s_grupo
            cv.Fields("FECHA") = PrvFecorte
            cv.Fields("NUMES") = s_numes
            cv.Fields("DESCRIP") = "COMISION POR INVERSION"
            cv.Fields("REFERENC") = s_socio

            cv.Update
            APLICA_COMISION_INVERSION
        End If
        
    End With
        'If s_socio = "208" Then
        '    IntRespuesta = MsgBox(s_intganado & " INTERESES DEVENGADOS." & s_comision, 0)
        'End If

End Sub
Private Sub CORTE_MES()
    
    's_saldomesa = 0
   'Abre DMOVPR Para agregar CARGO POR INTERESES
    If s_numes = 1 Then
       s_fecha = "30/11/2011"
    End If
    If s_numes = 2 Then
        s_fecha = "31/12/2011"
    End If
    If s_numes = 3 Then
        s_fecha = "31/01/2012"
    End If
    If s_numes = 4 Then
        s_fecha = "29/02/2012"
    End If
    If s_numes = 5 Then
        s_fecha = "31/03/2012"
    End If
    If s_numes = 6 Then
        s_fecha = "30/04/2012"
    End If
    If s_numes = 7 Then
        s_fecha = "31/05/2012"
    End If
    If s_numes = 8 Then
        s_fecha = "30/06/2012"
    End If
    If s_numes = 9 Then
        s_fecha = "31/07/2012"
    End If
    If s_numes = 10 Then
        s_fecha = "31/08/2012"
    End If
    If s_numes = 11 Then
        s_fecha = "30/09/2012"
    End If
    If s_numes = 12 Then
        s_fecha = "31/10/2012"
    End If

BUSCA_SOCIO
s_tasa = PrvTasaPres
If s_tipo = "7" Or s_tipo = 9 Then
    Exit Sub
End If
'If s_socio = "80" Then
'    IntRespuesta = MsgBox("SCORTE=" & s_corte & "CALCULA SOCIO=" & s_socio & " s_pagos=" & s_pagos & " s_saldopres=" & s_saldopres & " cs_prestamo=" & cs_prestamo & " saldomesa=" & s_saldomesa & " s_numes=" & s_numes & flg_next & flg_61 & " Importe: " & cs_importe, 0)
'End If
'If s_socio <> "80" Then
 '   Exit Sub
'End If
If s_saldomesa > 0 Then
    If s_pagos < PrvPagoMin Or s_pagos = 0 Then
        If s_saldomesa > PrvPagoMin Then
            If pago_anticipado = 0 Then
                s_tasa = s_tasa + 1
            Else
                pago_anticipado = 0
            End If
             'idmoroso()
        End If
    End If
    's_saldomesa = s_saldopres


    'If mes_pre = s_numes Then
    '   s_saldomesa = s_saldopres - cs_prestamo
    'End If

'If s_socio = "01" Then
'    IntRespuesta = MsgBox("SCORTE=" & s_corte & "CALCULA SOCIO=" & s_socio & " s_pagos=" & s_pagos & " s_saldopres=" & s_saldopres & " cs_prestamo=" & cs_prestamo & " saldomesa=" & s_saldomesa & " s_numes=" & s_numes & flg_next & flg_61, 0)
'End If
    
    s_cargoint = s_saldomesa * s_tasa / 100

    s_saldopres = s_saldopres + s_cargoint
    x_intpagados = x_intpagados + s_cargoint

'If s_socio = "74" Then
'    IntRespuesta = MsgBox("CALCULA SOCIO=" & s_socio & " s_pagos=" & s_pagos & " s_saldopres=" & s_saldopres & " saldomesa=" & s_saldomesa & " cargoint=" & s_cargoint, 0)
'End If
    
    'cs_prestamo = 0
 If s_cargoint > 0.1 Then
  If s_numes > 0 Then
   If s_socio > "00" Then
  
    'IntRespuesta = MsgBox("CALCULA SOCIO=" & s_socio & " s_cargoint=" & s_cargoint & " s_saldopres=" & s_saldopres, 0)

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cp As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVPR

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
 Set cp = New ADODB.Recordset

    With cp
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVPR"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

    With cp
        .AddNew
        cp.Fields("NUMREG") = cp.Fields("Id")
        cp.Fields("TIPO") = "C"
        cp.Fields("APREPAC") = "C"
        cp.Fields("CVEMOV") = "66"
        cp.Fields("IMPORTE") = s_cargoint
        cp.Fields("SOCIO") = s_socio
        cp.Fields("GRUPO") = s_grupo
        cp.Fields("FECHA") = s_fecha
        cp.Fields("NUMES") = s_numes
        cp.Fields("DESCRIP") = "CARGO POR INTERESES"
        cp.Fields("REFERENC") = ""
        cp.Fields("TASA") = s_tasa
        cp.Update
    End With

   
    If s_pagos = 0 Then
         With cp
            .AddNew
            cp.Fields("NUMREG") = cp.Fields("Id")
            cp.Fields("TIPO") = "C"
            cp.Fields("APREPAC") = "P"
            cp.Fields("CVEMOV") = "52"
            cp.Fields("IMPORTE") = 0
            cp.Fields("SOCIO") = s_socio
            cp.Fields("GRUPO") = s_grupo
            cp.Fields("FECHA") = s_fecha
            cp.Fields("NUMES") = s_numes
            cp.Fields("DESCRIP") = "NO PAGO MENSUAL"
            cp.Fields("REFERENC") = ""
            cp.Fields("TASA") = s_tasa
            cp.Update
         End With
    End If
   End If
  End If
End If
End If

    APLICA_INTERESES_SOCIO

If s_cargoint > 0.01 Then
        
    If s_tipo = "3" And s_tasa > 2 Then
    'If s_socio = "01" Then
    '    IntRespuesta = MsgBox("PROMOTOR" & s_promotor & " CARGOINT=$" & s_cargoint, 0)
    'End If
        APLICA_COMISION_PROMOTOR
    End If
    'If s_socio = "250" Then
    '   IntRespuesta = MsgBox("SOCIO=" & s_socio & " PROMOTOR=" & s_promotor & "INTPAGADO" & s_cargoint, 0)
    'End If
End If
's_pagos = 0
no_pago = "0"
s_cargoint = 0

's_saldopres = 0
End Sub
Sub BUSCA_SOCIO()

    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."

   cl.MoveFirst

    Do Until cl.EOF = True
        'If cl.Fields("SOCIO") = "99" Then
        '    LblMesAnterior = cl.Fields("SALDO")
        'End If
        
        If cl.Fields("SOCIO") = "988" Then
            IntsInv = cl.Fields("SALDO")
        End If
        
        If cl.Fields("SOCIO") = s_socio Then
            PrvNombre = cl.Fields("NOMBRE")
            PrvPagoMin = cl.Fields("PAGOMIN")
            PrvTasaPres = cl.Fields("TASAPRES")
            s_promotor = cl.Fields("PROMOTOR")
            s_tipo = cl.Fields("TIPO")
            s_grupo = cl.Fields("GRUPO")
            s_saldoActual = cl.Fields("SALDO")
            'Exit Do
        End If
        cl.MoveNext
        Loop
        
End Sub
Private Sub APLICA_INTERESES_SOCIO()

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   Dim strPath As String

    strPath = "C:\" & Carpeta & "\SISFED.mdb"

    Set cn = New ADODB.Connection

    With cn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM SOCIOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = s_socio Then
            'If s_socio = "04" Then
            '    IntRespuesta = MsgBox("SOCIO=" & s_socio & " PROMOTOR=" & s_promotor & "INTPAGADO" & s_cargoint, 0)
            'End If
            cl.Fields("SALDOPRES") = cl.Fields("SALDOPRES") + s_cargoint
            cl.Fields("INTPAGADO") = cl.Fields("INTPAGADO") + s_cargoint
            If s_numes = MesCorte Then
                cl.Fields("ULTPAGO") = s_pagos
            End If
            cl.Update
            Exit Do
        End If
        cl.MoveNext
    Loop
t_prestamos = 0
End Sub
Private Sub APLICA_COMISION_PROMOTOR()
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   Dim strPath As String

    strPath = "C:\" & Carpeta & "\SISFED.mdb"

    Set cn = New ADODB.Connection

    With cn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM SOCIOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
If s_tasa > 2 Then
    s_comision = (s_cargoint / s_tasa) * (s_tasa - 2)
    'If s_socio = "203" Then
    '    IntRespuesta = MsgBox("SOCIO=" & s_socio & " PROMOTOR=" & s_promotor & " COMISION" & s_comision, 0)
    'End If
    cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = s_promotor Then
            cl.Fields("COMISION") = cl.Fields("COMISION") + s_comision
            s_grupo = cl.Fields("GRUPO")
            CREA_ABONO_POR_COMISION
            cl.Update
            Exit Do
        End If
        cl.MoveNext
    Loop
End If
s_cargoint = 0
End Sub
Private Sub APLICA_COMISION_INVERSION()
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   Dim strPath As String

    strPath = "C:\" & Carpeta & "\SISFED.mdb"

    Set cn = New ADODB.Connection

    With cn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM SOCIOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

    cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = s_promotor Then
            cl.Fields("COMISION") = cl.Fields("COMISION") + s_comision
            s_grupo = cl.Fields("GRUPO")
            cl.Update
            Exit Do
        End If
        cl.MoveNext
    Loop
s_cargoint = 0
End Sub
Private Sub CREA_ABONO_POR_COMISION()
    
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cv As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVIN

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
 Set cv = New ADODB.Recordset

    With cv
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVIN"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With

    With cv
     If s_comision > 0 Then
        .AddNew
        cv.Fields("NUMREG") = cv.Fields("Id")
        cv.Fields("TIPO") = "T"
        cv.Fields("APREPAC") = "A"
        cv.Fields("CVEMOV") = "98"
        cv.Fields("IMPORTE") = s_comision
        cv.Fields("SOCIO") = s_promotor
        cv.Fields("GRUPO") = s_grupo
        cv.Fields("FECHA") = s_fecha
        cv.Fields("NUMES") = s_numes
        cv.Fields("DESCRIP") = "ABONO POR COMISION"
        cv.Fields("REFERENC") = s_socio
        cv.Fields("TASA") = s_tasa
        cv.Update
     End If
    End With
    
    TotComision = TotComision + s_comision
End Sub
Private Sub GRABA_SOCIO_MOROSO()
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   Dim strPath As String

    strPath = "C:\" & Carpeta & "\SISFED.mdb"

    Set cn = New ADODB.Connection

    With cn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM SOCIOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = s_socio Then
            If MesCorte = 8 Then
                cl.Fields("AGO") = "A"
                cl.Update
                Exit Do
            End If
        End If
        cl.MoveNext
    Loop
End Sub
Private Sub CREA_INI()

'ABRE UPDATE DETMOV

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cv As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVIN
   Dim cp As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVPR

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
Set cv = New ADODB.Recordset

    With cv
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVIN"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    With cv
        .AddNew
        cv.Fields("TIPO") = "A"
        cv.Fields("APREPAC") = "A"
        cv.Fields("CVEMOV") = "00"
        cv.Fields("IMPORTE") = s_invini
        cv.Fields("SOCIO") = s_socio
        cv.Fields("GRUPO") = s_grupo
        cv.Fields("FECHA") = TxtInicioEjercicio
        cv.Fields("NUMES") = 1
        cv.Fields("DESCRIP") = "SALDO ANTERIOR"
        cv.Fields("REFERENC") = ""
        cv.Update
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
    
    With cp
        .AddNew
        cp.Fields("TIPO") = "A"
        cp.Fields("APREPAC") = "C"
        cp.Fields("CVEMOV") = "00"
        cp.Fields("IMPORTE") = s_prestini
        cp.Fields("SOCIO") = s_socio
        cp.Fields("GRUPO") = s_grupo
        cp.Fields("FECHA") = TxtInicioEjercicio
        cp.Fields("NUMES") = 1
        cp.Fields("DESCRIP") = "SALDO ANTERIOR"
        cp.Fields("REFERENC") = ""
        cp.Update
    End With

End Sub
Private Sub BORRA_SOCIOS()
'IntRespuesta = MsgBox("BORRA SOCIOS", 0)

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
    
    'Borra DATOS del Archivo SOCIOS
    cl.MoveFirst

    Do Until cl.EOF = True
        cl.Fields("PROM_INV") = 0
        cl.Fields("INTPAGADO") = 0
        cl.Fields("COMISION") = 0
        cl.Fields("INTGANADO") = 0
        cl.Fields("ULTPAGO") = 0
        cl.Fields("PROM_INV") = 0
        cl.Fields("SALDOPRES") = cl.Fields("PRES_INI") + cl.Fields("PRESTAMOS") - cl.Fields("PAGOS")
        cl.Fields("FECORTE") = TxtFecorte
        cl.Update
        cl.MoveNext
    Loop
End Sub
Private Sub CREA_MOVIN()
'ABRE UPDATE DMOVIN

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cv As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVIN

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
Set cv = New ADODB.Recordset

    With cv
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVIN"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    With cv
        .AddNew
        cv.Fields("TIPO") = s_tipo
        cv.Fields("APREPAC") = s_aprepac
        cv.Fields("CVEMOV") = s_cvemov
        cv.Fields("IMPORTE") = s_importe
        cv.Fields("SOCIO") = s_socio
        cv.Fields("GRUPO") = s_grupo
        cv.Fields("FECHA") = s_fecha
        cv.Fields("NUMES") = s_numes
        cv.Fields("DESCRIP") = s_descrip
        cv.Fields("REFERENC") = s_referenc
        cv.Update
    End With
End Sub
Private Sub CREA_MOVPR()

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cp As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVPR

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
Set cp = New ADODB.Recordset

    With cp
        .ActiveConnection = cn
        .Source = "SELECT * FROM DMOVPR"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    With cp
        .AddNew
        cp.Fields("TIPO") = s_tipo
        cp.Fields("APREPAC") = s_aprepac
        cp.Fields("CVEMOV") = s_cvemov
        cp.Fields("IMPORTE") = s_importe
        cp.Fields("SOCIO") = s_socio
        cp.Fields("GRUPO") = s_grupo
        cp.Fields("FECHA") = s_fecha
        cp.Fields("NUMES") = s_numes
        cp.Fields("DESCRIP") = s_descrip
        cp.Fields("REFERENC") = s_referenc
        cp.Update
    End With
 

End Sub

Private Sub BORRA_DMOVPR()
'IntRespuesta = MsgBox("BORRA DMOVPR", 0)

    'Borra Registros de DMOVPR
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cp As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVPR

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
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
    'If Not cp.EOF Then
    cp.MoveFirst
    'End If
    'LblProceso.Caption = "Borra registros de DMOVPR"

    Do Until cp.EOF = True
        'IntRespuesta = MsgBox("CVEMOV=" & cp.Fields("CVEMOV") & "TIPO=" & cp.Fields("TIPO"), 0)
        'Exit Do
        
        PrvNumovs = PrvNumovs + 1
        If cp.Fields("IMPORTE") = 0 Then
            'And cp.Fields("NUMES") > 1 Then
            If Not cp.EOF Then
                s_delete = 1
                cp.Delete
                cp.MoveNext
            End If
        End If
        If cp.Fields("DESCRIP") = "CARGO POR INTERESES" Then
            'And cp.Fields("NUMES") > 1 Then
            If Not cp.EOF Then
                s_delete = 1
                cp.Delete
                cp.MoveNext
            End If
        End If
        If Not cp.EOF Then
            If cp.Fields("DESCRIP") = "NO PAGO MENSUAL" Then
                s_delete = 1
                cp.Delete
                cp.MoveNext
            End If
        End If
        If s_delete = 1 Then
            s_delete = 0
        Else
            cp.MoveNext
        End If
    Loop
End Sub
Private Sub BORRA_DETMOV()
'IntRespuesta = MsgBox("BORRA DMOVPR", 0)

    'Borra Registros de DMOVPR
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el Objeto Recordset.DETMOV

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
Set cd = New ADODB.Recordset

    With cd
        .ActiveConnection = cn
        .Source = "SELECT * FROM DETMOV"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    cd.MoveFirst

    Do Until cd.EOF = True
        'IntRespuesta = MsgBox("CVEMOV=" & cp.Fields("CVEMOV") & "TIPO=" & cp.Fields("TIPO"), 0)
        'Exit Do
        
        PrvNumovs = PrvNumovs + 1
        If cd.Fields("DESCRIP") = "ABONO POR COMISION" Then
            If Not cd.EOF Then
                s_delete = 1
                cd.Delete
                cd.MoveNext
            End If
        End If
        If cd.Fields("DESCRIP") = "CARGO POR INTERESES" Then
            If Not cd.EOF Then
                s_delete = 1
                cd.Delete
                cd.MoveNext
            End If
        End If
        If Not cd.EOF Then
            If cd.Fields("DESCRIP") = "NO PAGO MENSUAL" Then
                s_delete = 1
                cd.Delete
                cd.MoveNext
            End If
        End If
        If cd.Fields("SOCIO") = "" Then
            If Not cd.EOF Then
                s_delete = 1
                cd.Delete
                cd.MoveNext
            End If
        End If
        If s_delete = 1 Then
            s_delete = 0
        Else
            cd.MoveNext
        End If
    Loop
End Sub
Private Sub BORRA_DMOVIN()
'IntRespuesta = MsgBox("BORRA DMOVIN", 0)

    'Borra Registros de DMOVPR
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cv As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVPR

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
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
        If cv.Fields("DESCRIP") = "COMISION POR INVERSION" Then
            If Not cv.EOF Then
                s_delete = 1
                cv.Delete
                cv.MoveNext
            End If
        End If
        
        If cv.Fields("SOCIO") = "   " Then
            If Not cv.EOF Then
                s_delete = 1
                cv.Delete
                cv.MoveNext
            End If
        End If
        'And cv.Fields("TIPO") <> "D" Then
        
        If cv.Fields("DESCRIP") = "ABONO POR COMISION" Then
            If Not cv.EOF Then
                s_delete = 1
                cv.Delete
                cv.MoveNext
            End If
        End If
                
        If cv.Fields("DESCRIP") = "INTERESES DEVENGADOS" Then
            If Not cv.EOF Then
                s_delete = 1
                cv.Delete
                cv.MoveNext
            End If
        End If
        
        If s_delete = 1 Then
            s_delete = 0
        Else
            cv.MoveNext
        End If
    Loop
End Sub
Private Sub CREA_PROMEDIOS()
IntRespuesta = MsgBox("CREA_PROMEDIOS", 0)

Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   
   Dim cv As New ADODB.Recordset 'Creamos el Objeto Recordset.DMOVIN

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
    cv.Open "SELECT * FROM DMOVIN ORDER BY SOCIO,FECHA,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    cv.MoveFirst
    s_socio = cv.Fields("SOCIO")
    s_numes = Month(PrvFecorte)
    s_numes = s_numes + 2
    'LblProceso.Caption = "Inicia Cálculo de Intereses sobre Préstamos"
  
    Do Until cv.EOF = True
    x = x + 3
    Me.CurrentX = x
    Me.CurrentY = 6
    Me.Print ">>>"
    Me.Print "Espere un Momento Por Favor..."
    
       If cv.Fields("GRUPO") <> "99" Then
'IntRespuesta = MsgBox(cv.Fields("SOCIO") & ".-" & s_socio & " S_PROMINV=" & _
        s_prominv & " IMPORTE=" & cv.Fields("IMPORTE"), 0)

       If cv.Fields("SOCIO") <> "   " Then
        If cv.Fields("SOCIO") = s_socio Then

            If cv.Fields("APREPAC") = "A" Then

                s_prominv = cv.Fields("IMPORTE") * (diad - cv.Fields("FECHA"))
                If cv.Fields("CVEMOV") <> "00" Then
                    s_aporta = s_aporta + cv.Fields("IMPORTE")
                End If

            End If
            If cv.Fields("APREPAC") = "R" Then
                s_prominv = cv.Fields("IMPORTE") * (diad - cv.Fields("FECHA"))
                s_prominv = s_prominv * -1
                s_retiros = s_retiros + cv.Fields("IMPORTE")
            End If
            'IntRespuesta = MsgBox(cv.Fields("DESCRIP") & ".-" & s_socio & " S_PROMINV=" & s_prominv & " IMPORTE=" & cv.Fields("IMPORTE"), 0)
            t_prominv = t_prominv + s_prominv
            cv.MoveNext
        Else
'If s_socio = "997" Then
'    IntRespuesta = MsgBox("APREPAC=" & cv.Fields("APREPAC") & " socio.-" & s_socio & _
'        "  " & t_prominv & "  " & ac_prominv, 0)
'End If
            If s_socio > "00" Then
                t_prominv = t_prominv / 365
                ac_prominv = ac_prominv + t_prominv

                GRABA_PROMEDIO
            End If
            s_socio = cv.Fields("SOCIO")
            t_prominv = 0
            s_aporta = 0
            s_retiros = 0
        End If
       End If
       
    Else
        cv.MoveNext
    End If
    Loop
    If s_socio > "00" Then
        t_prominv = t_prominv / 365
        ac_prominv = ac_prominv + t_prominv
        GRABA_PROMEDIO
    End If
End Sub
Private Sub GRABA_PROMEDIO()
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"

   Dim strPath As String

    strPath = "C:\" & Carpeta & "\SISFED.mdb"

    Set cn = New ADODB.Connection

    With cn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source").Value = strPath
        .Open
    End With

Set cl = New ADODB.Recordset

    With cl
        .ActiveConnection = cn
        .Source = "SELECT * FROM SOCIOS"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseServer
        .Open
    End With
    
    cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = s_socio Then
            If cl.Fields("INV_INI") < 0.01 And cl.Fields("APORTA") < 0.01 Then
                cl.Fields("PROM_INV") = 0
            Else
                cl.Fields("PROM_INV") = t_prominv
            End If
            'IntRespuesta = MsgBox("GRABA PROM_INV=" & s_socio & t_prominv, 0)

            cl.Fields("PROM_APORT") = cl.Fields("APORTA") / s_numes
            If MesCorte = 9 Then
                cl.Fields("SEP") = ""
            End If
            cl.Update
            Exit Do
        End If
        cl.MoveNext
    Loop
  
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MS 
   Caption         =   "FrmGrid"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSocio 
      Height          =   2055
      Left            =   12360
      Picture         =   "FrmGrid.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   1875
      TabIndex        =   20
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdConsultaSocio 
      Caption         =   "CONSULTA SOCIO"
      Height          =   495
      Left            =   6240
      TabIndex        =   18
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox TxtCapRet 
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
      Left            =   2280
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtCapPres 
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
      Left            =   2280
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtGrpMS 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton CmdListGrp 
      Caption         =   "Lista de Socios por Grupo"
      Height          =   495
      Left            =   8040
      TabIndex        =   12
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton CmdNuevos 
      Caption         =   "Busca Claves para Nuevos Socios"
      Height          =   495
      Left            =   9840
      TabIndex        =   9
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox TxtNumMS 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton CmdSocNum 
      Caption         =   "Lista de Socios por Número"
      Height          =   615
      Left            =   12600
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox TxtNomMS 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton CmdCeros 
      Caption         =   "Lista Socios en Ceros"
      Height          =   615
      Left            =   12600
      TabIndex        =   8
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdPrestamos 
      Caption         =   "Lista Socios por Saldo de Préstamo"
      Height          =   615
      Left            =   12600
      TabIndex        =   6
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton CmdSaldos 
      Caption         =   "Lista Socios por Saldo de Inversión"
      Height          =   615
      Left            =   12600
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton CmdNombres 
      Caption         =   "Lista Socios por Nombre"
      Height          =   615
      Left            =   12600
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MS 
      Bindings        =   "FrmGrid.frx":1577
      Height          =   5655
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9975
      _Version        =   393216
      Rows            =   2000
      Cols            =   9
      ScrollTrack     =   -1  'True
   End
   Begin VB.Label LblCorte 
      Caption         =   "Corte al 30/Sep/2011"
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
      Left            =   9960
      TabIndex        =   21
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label LblEncab 
      Caption         =   "RELACION DE SOCIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   360
      Width           =   8055
   End
   Begin VB.Label Label5 
      Caption         =   "Capacidad de Préstamos"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Capacidad de Retiros"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Selecciona Grupo"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Selecciona Socio"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Selecciona por Nombre"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   6960
      Width           =   3015
   End
End
Attribute VB_Name = "MS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Carpeta As String
Public PubSocio As String
Private FlgMS As Integer
 'EJMPLO DE SUSO DEL CONTROL MS FLEXGRID
'SE TRATA DE UNA BASE DE DATOS (ACCESS(MDB))  LLAMADA  SISFED  QUE RESIDE EN EL DISCO C,
'CON UNA TABLA QUE SE LLAMA SOCIOS :
'QUE SU UBICACIÓN SERÍA C:\" & Carpeta & "\SISFED.MDB
'EL MSFLEX GRID SE LAMA SOLAMENTE MS
'PROCCEDIMIENTO PARA COLOCAR TITULOS  LAS COLUMNAS DEL GRID.'
Private totInvGrp, RenBorra, flgsocio As Single
Private totPresGrp As Single
Private totComGrp As Single
Private totIntGrp As Single

        'IntRespuesta = MsgBox(totInvGrp, 0)

Private Sub ValorFlexGrid()


   'IntRespuesta = MsgBox(MS.Text, 0)
End Sub
    
Sub COLOCATITULOSENMS()
MS.Row = 0
MS.Col = 0
MS.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
MS.Col = 1
MS.Text = "GRUPO"
MS.Col = 2
MS.Text = "SOCIO"
MS.Col = 3
MS.Text = "NOMBRE"
MS.Col = 4
MS.Text = "SALDO"
MS.Col = 5
MS.Text = "APORTA"
MS.Col = 6
MS.Text = "RETIROS"
MS.Col = 7
MS.Text = "INTGANADO"
MS.Col = 8
MS.Text = "COMISION"

MS.ColWidth(3) = 3000    'AJUSTO EL ANCHO DE LA COLUMNA2
MS.ColWidth(4) = 1000    'AJUSTO EL ANCHO DE LA COLUMNA2
MS.ColWidth(5) = 1000    'AJUSTO EL ANCHO DE LA COLUMNA2
Label5.Visible = False
Label4.Visible = False
TxtCapPres.Visible = False
TxtCapRet.Visible = False
TxtCapPres = 0
TxtCapRet = 0
totInvGrp = 0
totPresGrp = 0
totComGrp = 0
totIntGrp = 0
End Sub
Sub BorraCeldasenMS()
    RENGLON = RenBorra
    Do Until RENGLON = 199
       RENGLON = RENGLON + 1
       MS.Col = 0
       MS.Row = RENGLON
       MS.Text = ""
       MS.Col = 1
       MS.Text = ""
       MS.Col = 2
       MS.Text = ""
       MS.Col = 3
       MS.Text = ""
       MS.Col = 4
       MS.Text = ""
       MS.Col = 5
       MS.Text = ""
       MS.Col = 6
       MS.Text = ""
       MS.Col = 7
       MS.Text = ""
       MS.Col = 8
       MS.Text = ""
    Loop

End Sub
'PROCEDIMIENTO PARA OLOCAR LOS DATOS EN EL MS.
Sub COLOCADATOSENMS()


   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SOCIOS ORDER BY NOMBRE", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
   
    Do Until cl.EOF = True
    RENGLON = RENGLON + 1
    MS.Col = 0
    MS.Row = RENGLON
    MS.Text = RENGLON
    MS.Col = 1
    MS.Text = ""
    MS.Col = 2
    MS.Text = cl.Fields("SOCIO")
    MS.Col = 3
    MS.Text = cl.Fields("NOMBRE")
    MS.Col = 4
    MS.Text = Format(cl.Fields("SALDO"), "Currency")
    MS.Col = 5
    MS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
            MS.Col = 6
            MS.Text = Format(cl.Fields("INTGANADO"), "Currency")
            MS.Col = 7
            MS.Text = Format(cl.Fields("COMISION"), "Currency")
    MS.Col = 8
    MS.Text = ""

    cl.MoveNext
Loop
cl.Close

End Sub
   Sub COLOCASALDOSENMS()
    Dim TotSaldo As Single

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SOCIOS ORDER BY SALDO DESC", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
   
    Do Until cl.EOF = True
        TotSaldo = TotSaldo + cl.Fields("SALDO")
        cl.MoveNext
        Loop
    cl.MoveFirst
    Do Until cl.EOF = True
    If cl.Fields("SALDO") > 0 Then
    If cl.Fields("TIPO") <> "9" Then

        Porciento = cl.Fields("SALDO") / TotSaldo
        TotPorciento = TotPorciento + Porciento
        RENGLON = RENGLON + 1
        MS.Col = 0
        MS.Row = RENGLON
        MS.Text = RENGLON
        MS.Col = 1
        MS.Text = Format(Porciento, "Percent")
        MS.Col = 2
        MS.Text = cl.Fields("SOCIO")
        MS.Col = 3
        MS.Text = cl.Fields("NOMBRE")
        MS.Col = 4
        MS.Text = Format(cl.Fields("SALDO"), "Currency")
        MS.Col = 5
        MS.Text = Format(cl.Fields("APORTA"), "Currency")
        MS.Col = 6
        MS.Text = Format(cl.Fields("RETIROS"), "Currency")
        MS.Col = 7
        MS.Text = Format(cl.Fields("INTGANADO"), "Currency")
        MS.Col = 8
        MS.Text = Format(cl.Fields("COMISION"), "Currency")

        If TotPorciento > 0.5 Then
            RENGLON = RENGLON + 1
            MS.Row = RENGLON
            MS.Col = 1
            MS.Text = Format(TotPorciento, "Percent")
            MS.Col = 2
            MS.Text = ""
            TotPorciento = 0
            MS.Col = 3
            MS.Text = "SOCIOS MAYORITARIOS"
            MS.Col = 4
            MS.Text = ""
            MS.Col = 5
            MS.Text = ""
            MS.Col = 6
            MS.Text = ""
            MS.Col = 7
            MS.Text = ""
            MS.Col = 8
            MS.Text = ""
            RENGLON = RENGLON + 1
            MS.Row = RENGLON
            MS.Col = 1
            MS.Text = ""
            MS.Col = 2
            MS.Text = ""
            MS.Col = 3
            MS.Text = ""
            MS.Col = 4
            MS.Text = ""
            MS.Col = 5
            MS.Text = ""
            MS.Col = 6
            MS.Text = ""
            MS.Col = 7
            MS.Text = ""
            MS.Col = 8
            MS.Text = ""
        End If
        'IntRespuesta = MsgBox(TotSaldo, 0)

    End If
    End If
    cl.MoveNext
Loop
RenBorra = RENGLON
BorraCeldasenMS
cl.Close

End Sub


Sub COLOCAPRESTAMOENMS()
MS.Row = 0
MS.Col = 0
MS.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
MS.Col = 1
MS.Text = "GRUPO"
MS.Col = 2
MS.Text = "SOCIO"
MS.Col = 3
MS.Text = "NOMBRE"
MS.Col = 4
MS.Text = "PRESTAMO"
MS.Col = 5
MS.Text = "PAGOTOT"
MS.Col = 6
MS.Text = "PAGOMIN"
MS.Col = 7
MS.Text = "TASA INT"
MS.Col = 8
MS.Text = "ULT PAGO"

MS.ColWidth(1) = 850    'AJUSTO EL ANCHO DE LA COLUMNA2
MS.ColWidth(2) = 850    'AJUSTO EL ANCHO DE LA COLUMNA2
MS.ColWidth(3) = 3000    'AJUSTO EL ANCHO DE LA COLUMNA2
Label5.Visible = False
Label4.Visible = False

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SOCIOS ORDER BY SALDOPRES, ULTPAGO, GRUPO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
   
    Do Until cl.EOF = True
        If cl.Fields("TIPO") <> "9" Then
          If cl.Fields("SALDOPRES") <> 0 Then
            RENGLON = RENGLON + 1
            MS.Col = 0
            MS.Row = RENGLON
            MS.Text = RENGLON
            MS.Col = 1
            MS.Text = cl.Fields("GRUPO")
            MS.Col = 2
            MS.Text = cl.Fields("SOCIO")
            MS.Col = 3
            MS.Text = cl.Fields("NOMBRE")
            MS.Col = 4
            MS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
            MS.Col = 5
            MS.Text = Format(cl.Fields("SALDOPRES") * cl.Fields("TASAPRES") / 100 + cl.Fields("SALDOPRES"), "Currency")
            MS.Col = 6
            MS.Text = Format(cl.Fields("PAGOMIN"), "Currency")
            MS.Col = 7
            MS.Text = Format(cl.Fields("TASAPRES") / 100, "Percent")
            MS.Col = 8
            MS.Text = Format(cl.Fields("ULTPAGO"), "Currency")
          End If
        End If
    cl.MoveNext
Loop
RenBorra = RENGLON
BorraCeldasenMS
cl.Close

End Sub
Sub BuscaSociosEnCeros()
    MS.ColWidth(2) = 1200    'AJUSTO EL ANCHO DE LA COLUMNA2
    MS.ColWidth(3) = 3000    'AJUSTO EL ANCHO DE LA COLUMNA2
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
   
    Do Until cl.EOF = True
        If cl.Fields("SALDOPRES") < 0.01 And cl.Fields("SALDO") < 0.01 And cl.Fields("APORTA") < 0.01 Then
         If cl.Fields("RETIROS") < 0.01 And cl.Fields("PRESTAMOS") < 0.01 And cl.Fields("PAGOS") < 0.01 Then
         

            RENGLON = RENGLON + 1
            MS.Col = 0
            MS.Row = RENGLON
            MS.Text = RENGLON
            MS.Col = 1
            MS.Text = ""
            MS.Col = 2
            MS.Text = cl.Fields("SOCIO")
            MS.Col = 3
            MS.Text = cl.Fields("NOMBRE")
            MS.Col = 4
            MS.Text = Format(cl.Fields("SALDO"), "Currency")
            MS.Col = 5
            MS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
            MS.Col = 6
            MS.Text = ""
            MS.Col = 7
            MS.Text = ""
            MS.Col = 8
            MS.Text = ""

         End If
        End If
    cl.MoveNext
Loop
RenBorra = RENGLON
BorraCeldasenMS
cl.Close

End Sub

Private Sub CmdCeros_Click()
COLOCATITULOSENMS
'BorraCeldasenMS

BuscaSociosEnCeros
End Sub

Private Sub CmdConsultaSocio_Click()

  Static lfrmCount As Long
    Dim frmD As FrmSocios
    lfrmCount = lfrmCount + 1
    Set frmD = New FrmSocios
    frmD.Caption = "frmSocios"
    
    frmD.Show
End Sub

Private Sub CmdListGrp_Click()
COLOCATITULOSENMS
MS.Col = 6
MS.Text = "INTGANADO"
MS.Col = 7
MS.Text = "COMISION"
MS.Col = 8
MS.Text = ""
'COLOCADATOSENMS
Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SOCIOS ORDER BY GRUPO,SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
   sgrupo = cl.Fields("GRUPO")
    Do Until cl.EOF = True
            If sgrupo <> cl.Fields("GRUPO") Then
                RENGLON = RENGLON + 1
                MS.Row = RENGLON
                MS.Col = 0
                MS.Text = RENGLON
                MS.Col = 1
                MS.Text = ""
                MS.Col = 2
                MS.Text = ""
                MS.Col = 3
                MS.Text = ""
                MS.Col = 4
                MS.Text = ""
                MS.Col = 5
                MS.Text = ""
                MS.Col = 6
                MS.Text = ""
                MS.Col = 7
                MS.Text = ""
                MS.Col = 8
                MS.Text = ""
                sgrupo = cl.Fields("GRUPO")
            End If
            RENGLON = RENGLON + 1
            MS.Col = 0
            MS.Row = RENGLON
            MS.Text = RENGLON
            MS.Col = 1
            MS.Text = cl.Fields("GRUPO")
            MS.Col = 2
            MS.Text = cl.Fields("SOCIO")
            MS.Col = 3
            MS.Text = cl.Fields("NOMBRE")
            MS.Col = 4
            MS.Text = Format(cl.Fields("SALDO"), "Currency")
            MS.Col = 5
            MS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
                        MS.Col = 6
            MS.Text = Format(cl.Fields("INTGANADO"), "Currency")
            MS.Col = 7
            MS.Text = Format(cl.Fields("COMISION"), "Currency")
            MS.Col = 8
            MS.Text = ""

    cl.MoveNext
Loop
cl.Close

End Sub

Private Sub CmdNombres_Click()
frmMiPrimera.Flg = 4
Unload Me
Static lfrmCount As Long
    Dim frmD As MS
    lfrmCount = lfrmCount + 1
    Set frmD = New MS
    frmD.Caption = "MS"

    frmD.Show
    FlgMS = 1
End Sub

Private Sub CmdNuevos_Click()
COLOCATITULOSENMS
'BorraCeldasenMS
Dim Clave_ant As String
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
    Clave_ant = 201

   
    cl.MoveFirst
    Do Until cl.EOF = True
          
       If cl.Fields("SOCIO") > Clave_ant Then
            'IntRespuesta = MsgBox(Clave_ant & " " & cl.Fields("SOCIO"), 0)

            RENGLON = RENGLON + 1
            MS.Col = 0
            MS.Row = RENGLON
            MS.Text = ""
            MS.Col = 1
            MS.Text = ""
            MS.Col = 2
            MS.Text = Clave_ant
            MS.Col = 3
            MS.Text = ""
            
            MS.Col = 4
            MS.Text = ""
            MS.Col = 5
            MS.Text = ""
            MS.Col = 6
            MS.Text = ""
            MS.Col = 7
            MS.Text = ""
            Clave_ant = Clave_ant + 1
            'IntRespuesta = MsgBox(Clave_ant & " " & cl.Fields("SOCIO"), 0)
        End If
        If cl.Fields("SOCIO") = Clave_ant Then
            cl.MoveNext
            Clave_ant = Clave_ant + 1
        End If
        If cl.Fields("SOCIO") < Clave_ant Then
            cl.MoveNext
        End If

        If RENGLON = 10 Then
            Exit Do
        End If
             

    Loop
RenBorra = RENGLON
BorraCeldasenMS
cl.Close

End Sub

Private Sub CmdPrestamos_Click()
frmMiPrimera.Flg = "2"
Unload Me
Static lfrmCount As Long
    Dim frmD As MS
    lfrmCount = lfrmCount + 1
    Set frmD = New MS
    frmD.Caption = "MS"

    frmD.Show
    'FlgMS = 1
    'MsgBox ("frmMiPrimera=" & frmMiPrimera.Flg)
End Sub
'COLOCATITULOSENMS

'BorraCeldasenMS
'C'OLOCAPRESTAMOENMS

'End Sub

Private Sub CmdSaldos_Click()
frmMiPrimera.Flg = 1
Unload Me
Static lfrmCount As Long
    Dim frmD As MS
    lfrmCount = lfrmCount + 1
    Set frmD = New MS
    frmD.Caption = "MS"

    frmD.Show
    FlgMS = 1
End Sub

'COLOCATITULOSENMS
'BorraCeldasenMS
'COLOCASALDOSENMS
'End Sub


Private Sub CmdSocNum_Click()
frmMiPrimera.Flg = "3"
Unload Me
Static lfrmCount As Long
    Dim frmD As MS
    lfrmCount = lfrmCount + 1
    Set frmD = New MS
    frmD.Caption = "MS"

    frmD.Show
End Sub

Private Sub SociosPorNumero()
MS.ColWidth(2) = 3000    'AJUSTO EL ANCHO DE LA COLUMNA2
MS.ColWidth(3) = 1200    'AJUSTO EL ANCHO DE LA COLUMNA2
MS.ColWidth(4) = 1200    'AJUSTO EL ANCHO DE LA COLUMNA2
MS.ColWidth(5) = 1200    'AJUSTO EL ANCHO DE LA COLUMNA2

'COLOCATITULOSENMS
MS.Row = 0
MS.Col = 0
MS.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
MS.Col = 1
MS.Text = "SOCIO"
MS.Col = 2
MS.Text = "NOMBRE"
MS.Col = 3
MS.Text = "INVERSION"
MS.Col = 4
MS.Text = "APORTACION"
MS.Col = 5
MS.Text = "RETIRO"
MS.Col = 6
MS.Text = "INT_GANADO"
MS.Col = 7
MS.Text = "COMISION"
MS.Col = 8
MS.Text = "PROMEDIO"
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
   
    Do Until cl.EOF = True
            t_saldo = t_saldo + cl.Fields("SALDO")
            t_saldopres = t_saldopres + cl.Fields("SALDOPRES")
            t_prominv = t_prominv + cl.Fields("PROM_INV")
            t_intganado = t_intganado + cl.Fields("INTGANADO")
            t_comision = t_comision + cl.Fields("COMISION")
            t_intpagado = t_intpagado + cl.Fields("INTPAGADO")
            RENGLON = RENGLON + 1
            MS.Col = 0
            MS.Row = RENGLON
            MS.Text = RENGLON
            MS.Col = 1
            MS.Text = cl.Fields("SOCIO")
            MS.Col = 2
            MS.Text = cl.Fields("NOMBRE")
            MS.Col = 3
            MS.Text = Format(cl.Fields("SALDO"), "Currency")
            MS.Col = 4
            MS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
            MS.Col = 5
            MS.Text = Format(cl.Fields("PROM_INV"), "Currency")
            MS.Col = 6
            MS.Text = Format(cl.Fields("INTGANADO"), "Currency")
            MS.Col = 7
            MS.Text = Format(cl.Fields("COMISION"), "Currency")
            MS.Col = 8
            MS.Text = Format(cl.Fields("INTPAGADO"), "Currency")
    cl.MoveNext
Loop
            RENGLON = RENGLON + 1
            MS.Col = 0
            MS.Row = RENGLON
            MS.Text = RENGLON
            MS.Col = 3
            MS.Text = Format(t_saldo, "Currency")
            MS.Col = 4
            MS.Text = Format(t_saldopres, "Currency")
            MS.Col = 5
            MS.Text = Format(t_prominv, "Currency")
            MS.Col = 6
            MS.Text = Format(t_intganado, "Currency")
            MS.Col = 7
            MS.Text = Format(t_comision, "Currency")
            MS.Col = 8
            MS.Text = Format(t_intpagado, "Currency")
cl.Close

End Sub





Private Sub Form_Load()
    'IntRespuesta = MsgBox("Mi Primera Carpeta=" & frmMiPrimera.LblCarpeta, 0)
    Carpeta = frmMiPrimera.LblCarpeta
    'IntRespuesta = MsgBox("Carpeta=" & Carpeta, 0)
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb" '''
   cl.Open "SELECT * FROM SOCIOS ORDER BY GRUPO,SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   cl.MoveFirst
    LblCorte = "Corte al " & cl.Fields("FECORTE")

    PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Fedamac.jpg")

    If frmMiPrimera.Flg = "4" Then
        COLOCATITULOSENMS
        COLOCADATOSENMS
    End If

    If frmMiPrimera.Flg = "3" Then
        LblEncab.Caption = "LISTA DE SOCIOS POR NUMERO"
        SociosPorNumero
    End If
   If frmMiPrimera.Flg = "2" Then
        LblEncab.Caption = "RELACION DE SOCIOS POR PRÉSTAMOS"

        COLOCATITULOSENMS
        COLOCAPRESTAMOENMS
        'frmMiPrimera.Flg = "0"
    End If
   If frmMiPrimera.Flg = "1" Then
        LblEncab.Caption = "RELACION DE SOCIOS POR INVERSION"
        'MsgBox ("MS.LblTitulo=" & LblEncab.Caption)
        COLOCATITULOSENMS
        COLOCASALDOSENMS
    End If
   ' MsgBox ("frmMiprimera=" & frmMiPrimera.Flg)

    If frmMiPrimera.Flg = "0" Then
        LblEncab.Caption = "CAPACIDAD POR GRUPOS"

        TxtGrpMS = "02"
        TxtGrpMSEnter
    End If
End Sub

Private Sub MS_Click()
    'If MS.Col = 2 Then
        PubSocio = MS.Text
        frmMiPrimera.LblSocio = MS.Text
        'IntRespuesta = MsgBox(PubSocio, 0)
        'TxtclSocio.Text = MS.Text
        'Static lfrmCount As Long
        'Dim frmD As FrmSocios
        'lfrmCount = lfrmCount + 1
        'Set frmD = New FrmSocios
        'frmD.Caption = "frmSocios"
    
        'frmD.Show
    'End If
End Sub
Private Sub TxtGrpMS_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtGrpMS.Tag = MODE_OVERTYPE And TxtGrpMS.SelLength = 0 Then
        TxtGrpMS.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtGrpMSEnter
       'SendKeys "{tab}"
    End If
    If flgsocio <> 1 Then
        If KeyAscii <> 13 Then
            flgsocio = 1
            TxtGrpMS = ""
        End If
    End If
End Sub

'LISTA SOCIOS POR UN GRUPO ESPECIFICO
Private Sub TxtGrpMSEnter()
        PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Fedamac.jpg")
  
        If TxtGrpMS = "01" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\David.jpg")
        End If
        If TxtGrpMS = "02" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Luis2.jpg")
        End If
        If TxtGrpMS = "03" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Lola.jpg")
        End If
        If TxtGrpMS = "04" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Chuy.jpg")
        End If
        If TxtGrpMS = "05" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Yola.jpg")
        End If
        If TxtGrpMS = "06" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Julio.jpg")
        End If
        If TxtGrpMS = "07" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\JuanToña.jpg")
        End If
        If TxtGrpMS = "08" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\FerLilia.jpg")
        End If
        If TxtGrpMS = "09" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\RaulG.jpg")
        End If
        If TxtGrpMS = "10" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\RosaG.jpg")
        End If
        If TxtGrpMS = "11" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\FotoG11.jpg")
        End If

        If TxtGrpMS = "12" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Paty.jpg")
        End If
        If TxtGrpMS = "12" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Mario.jpg")
        End If

        If TxtGrpMS = "13" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Paty.jpg")
        End If
        If TxtGrpMS = "14" Then
            PicSocio.Picture = LoadPicture("C:\" & Carpeta & "\Alicia.jpg")
        End If
LblEncab.Caption = "CAPACIDAD POR GRUPOS"

Dim totInvGrp As Single
totInvGrp = 0
COLOCATITULOSENMS
MS.Col = 5
MS.Text = "PRESTAMOS"
MS.Col = 6
MS.Text = "INTGANADO"
MS.Col = 7
MS.Text = "COMISION"
MS.Col = 8
MS.Text = ""
'BorraCeldasenMS
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   cl.Open "SELECT * FROM SOCIOS ORDER BY GRUPO,SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   cl.MoveFirst
   Do Until cl.EOF = True

   If TxtGrpMS = cl.Fields("GRUPO") Then
        If flgsocio = 1 Then
            TxtNumMS = cl.Fields("SOCIO")
            frmMiPrimera.LblSocio = TxtNumMS

            flgsocio = 0
        End If
        RENGLON = RENGLON + 1
        MS.Col = 0
        MS.Row = RENGLON
        MS.Text = RENGLON
        MS.Col = 1
        MS.Text = cl.Fields("GRUPO")
        MS.Col = 2
        MS.Text = cl.Fields("SOCIO")
        MS.Col = 3
        MS.Text = cl.Fields("NOMBRE")
        MS.Col = 4
        MS.Text = Format(cl.Fields("SALDO"), "###,###,##0.00")

        totInvGrp = totInvGrp + cl.Fields("SALDO")
        MS.Col = 5
        MS.Text = Format(cl.Fields("SALDOPRES"), "###,###,##0.00")
        totPresGrp = totPresGrp + Format(cl.Fields("SALDOPRES"))
        MS.Col = 6
        MS.Text = Format(cl.Fields("INTGANADO"), "###,###,##0.00")
        totIntGrp = totIntGrp + Format(cl.Fields("INTGANADO"))
        MS.Col = 7
        MS.Text = Format(cl.Fields("COMISION"), "###,###,##0.00")
        totComGrp = totComGrp + Format(cl.Fields("COMISION"))
        MS.Col = 8
        MS.Text = ""
    End If
    cl.MoveNext
Loop
RENGLON = RENGLON + 1
MS.Col = 0
MS.Row = RENGLON
MS.Text = ""
MS.Col = 1
MS.Text = ""
MS.Col = 2
MS.Text = ""
RENGLON = RENGLON + 1
MS.Col = 3
MS.Text = "TOTAL POR GRUPO"
MS.Col = 4
MS.Text = Format(totInvGrp, "Currency")
MS.Col = 5
MS.Text = Format(totPresGrp, "Currency")
MS.Col = 6
MS.Text = Format(totIntGrp, "Currency")
MS.Col = 7
MS.Text = Format(totComGrp, "Currency")

Label5.Visible = True
Label4.Visible = True
TxtCapPres.Visible = True
TxtCapRet.Visible = True
TxtCapPres = Format((totInvGrp + totIntGrp + totComGrp) * 2 - totPresGrp, "Currency")
TxtCapRet = Format(TxtCapPres / 2, "Currency")

RenBorra = RENGLON - 1
BorraCeldasenMS
cl.Close
flgsocio = 0

End Sub

Private Sub TxtNomMS_Change()
COLOCATITULOSENMS
'BorraCeldasenMS
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'ESTA LINEA SOBRA porque ya has especificado el origen en la SELECT:
    'cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"

   'Esto tiene que funcionar
   cl.Open "SELECT * FROM SOCIOS ORDER BY NOMBRE", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
   cl.MoveFirst
    LongNom = 1
    LongNom = Len(TxtNomMS)
    Do Until cl.EOF = True
        varnombre = Left(cl.Fields("NOMBRE"), LongNom)
        UNOMBRE = TxtNomMS
        UNOMBRE = UCase(UNOMBRE)

        If varnombre = UNOMBRE Then
            RENGLON = RENGLON + 1
            MS.Col = 0
            MS.Row = RENGLON
            MS.Text = RENGLON
            MS.Col = 2
            MS.Text = cl.Fields("SOCIO")
            MS.Col = 3
            MS.Text = cl.Fields("NOMBRE")
            MS.Col = 4
            MS.Text = Format(cl.Fields("SALDO"), "Currency")
            MS.Col = 5
            MS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
                        MS.Col = 6
            MS.Text = Format(cl.Fields("INTGANADO"), "Currency")
            MS.Col = 7
            MS.Text = Format(cl.Fields("COMISION"), "Currency")


        End If
        cl.MoveNext
    Loop
RenBorra = RENGLON
BorraCeldasenMS
cl.Close
ValorFlexGrid

End Sub
Private Sub TxtNumMS_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtNumMS.Tag = MODE_OVERTYPE And TxtNumMS.SelLength = 0 Then
        TxtNumMS.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtNumMSEnter
       'SendKeys "{tab}"
    End If
    If flgsocio <> 1 Then
        If KeyAscii <> 13 Then
            flgsocio = 1
            TxtNumMS = ""
        End If
    End If
End Sub
'LISTA SOCIOS POR NUMERO
Private Sub TxtNumMSEnter()
PubSocio = TxtNumMS
frmMiPrimera.LblSocio = TxtNumMS

COLOCATITULOSENMS
'BorraCeldasenMS
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   cl.MoveFirst
   Do Until cl.EOF = True
If TxtNumMS <= cl.Fields("SOCIO") Then
    RENGLON = RENGLON + 1
    MS.Col = 0
    MS.Row = RENGLON
    MS.Text = RENGLON
    MS.Col = 1
    MS.Text = ""
    MS.Col = 2
    MS.Text = cl.Fields("SOCIO")
    MS.Col = 3
    MS.Text = cl.Fields("NOMBRE")
    MS.Col = 4
    MS.Text = Format(cl.Fields("SALDO"), "Currency")
    MS.Col = 5
    MS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
                MS.Col = 6
            MS.Text = Format(cl.Fields("INTGANADO"), "Currency")
            MS.Col = 7
            MS.Text = Format(cl.Fields("COMISION"), "Currency")


    End If
    cl.MoveNext
Loop
RenBorra = RENGLON
BorraCeldasenMS
cl.Close
flgsocio = 0

End Sub


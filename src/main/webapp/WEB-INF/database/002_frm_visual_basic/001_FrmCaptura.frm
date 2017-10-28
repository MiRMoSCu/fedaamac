VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCaptura 
   Caption         =   "CAPTURA MOVIMIENTOS"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtBanco 
      Height          =   285
      Left            =   9840
      TabIndex        =   32
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TxtSaldoCaja 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8520
      TabIndex        =   30
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TxtSocioFS 
      Height          =   285
      Left            =   1320
      TabIndex        =   27
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox TxtNomFs 
      Height          =   285
      Left            =   2280
      TabIndex        =   25
      Top             =   3120
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid FC 
      Height          =   1815
      Left            =   6120
      TabIndex        =   24
      Top             =   1320
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   40
      Cols            =   5
   End
   Begin MSFlexGridLib.MSFlexGrid FS 
      Height          =   1815
      Left            =   240
      TabIndex        =   23
      Top             =   1320
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   300
      Cols            =   6
      FocusRect       =   2
   End
   Begin VB.TextBox TxtUltNombre 
      Height          =   285
      Left            =   3960
      TabIndex        =   22
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox TxtUltSocio 
      Height          =   285
      Left            =   3360
      TabIndex        =   21
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtUltImporte 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """###,###,##0.00"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CmdUltima 
      Caption         =   "Ultima Transacción:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TxtAPrePac 
      Height          =   285
      Left            =   6840
      TabIndex        =   18
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox TxtTipo 
      Height          =   285
      Left            =   6600
      TabIndex        =   17
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton CmdGrabaReg 
      Caption         =   "GRABAR EL REGISTRO"
      Height          =   495
      Left            =   10800
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton CmdCveMov 
      Caption         =   "CVEMOV"
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton CmdSocio 
      Caption         =   "SOCIO"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox TxtReferenc 
      Height          =   285
      Left            =   10560
      MaxLength       =   10
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox TxtCtaBco 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9840
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "LLB"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox TxtDescrip 
      Height          =   285
      Left            =   7080
      MaxLength       =   23
      TabIndex        =   5
      Text            =   "DEPOSITO EN EFECTIVO"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox TxtCveMov 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6120
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "10"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtFecha 
      Height          =   285
      Left            =   5040
      TabIndex        =   3
      Text            =   "26/06/2010"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox TxtSocio 
      Height          =   285
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtImporte 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label TxtTasa 
      Height          =   375
      Left            =   12120
      TabIndex        =   31
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "SALDO EN CAJA:"
      Height          =   255
      Left            =   7080
      TabIndex        =   29
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Selecciona Socio"
      Height          =   375
      Left            =   1320
      TabIndex        =   28
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Selecciona por Nombre"
      Height          =   255
      Left            =   2280
      TabIndex        =   26
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "REFERENCIA"
      Height          =   255
      Left            =   10560
      TabIndex        =   15
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "BANCO"
      Height          =   255
      Left            =   9840
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "DESCRIPCION"
      Height          =   255
      Left            =   6960
      TabIndex        =   13
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "FECHA"
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "NOMBRE"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "IMPORTE"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FrmCaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Carpeta As String
Private dBanco As String
'Declare Function SetCursorPos Lib "user32.dll" _
() '(ByVal x As Long, ByVal y As Long) As Long
Private KeyAscii, flgsocio As Integer
Private SetCursorPos As String
Private n_socio, UBANCO, n_cvemov As String
Private prvImporte As String
Private PrvEnteros, s_numreg As Single
Private PrvSocio As String
Private PrvGrupo As String
Private Numovs, PrvTasa, TabIndex As Single
Private Clave As String
Private IntRespuesta As String
Private SaldoPres As Single
Private Saldo As Single
Private Fecha As String
Private CveMov As String
Private Tipo As String
Private Aprepac As String
Private CtaBco As String
Private Referenc As String
Private Provisional As String
Private RENGLON As Single
Private LongNom As Single
Private UNOMBRE As String
Private varnombre As String
Private TxtNomMS As String

Option Explicit

Const MODE_OVERTYPE = "overtype"
Const MODE_INSERT = "insert"


Private Sub Form_Load()
    Dim strPath
    flgsocio = 0
    'IntRespuesta = MsgBox("Mi Primera Carpeta=" & frmMiPrimera.LblCarpeta, 0)
    Carpeta = frmMiPrimera.LblCarpeta
    dBanco = frmMiPrimera.txtpsw
    'IntRespuesta = MsgBox("Carpeta=" & Carpeta & " DBanco " & dBanco, 0)
    strPath = "C:\" & Carpeta & "\" & dBanco & ".mdb"

    IntRespuesta = MsgBox("BASE DE DATOS: " & strPath, 0)

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb" '''
    
    TxtSocio.Tag = MODE_OVERTYPE
    TxtFecha.Tag = MODE_OVERTYPE
    TxtCveMov.Tag = MODE_OVERTYPE
    TxtDescrip.Tag = MODE_OVERTYPE
    TxtCtaBco.Tag = MODE_OVERTYPE
    TxtReferenc.Tag = MODE_OVERTYPE
    'Label1.Caption = MODE_INSERT
End Sub

Private Sub CmdCveMov_Click()
FC.Rows = 50
RENGLON = 0
FC.Row = 0
FC.Col = 0
FC.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
FC.Col = 1
FC.Text = "CVEMOV"
FC.Col = 2
FC.Text = "TIPO"
FC.Col = 3
FC.Text = "ARPC"
FC.Col = 4
FC.Text = "DESCRIPCION"

FC.ColWidth(0) = 450    'AJUSTO EL ANCHO DE LA COLUMNA

FC.ColWidth(1) = 700    'AJUSTO EL ANCHO DE LA COLUMNA
FC.ColWidth(2) = 200    'AJUSTO EL ANCHO DE LA COLUMNA2

FC.ColWidth(3) = 200    'AJUSTO EL ANCHO DE LA COLUMNA2
FC.ColWidth(4) = 2500    'AJUSTO EL ANCHO DE LA COLUMNA2

 'Sub BorraCeldasenMS()
    Do Until RENGLON = 49
        'MsgBox (RENGLON)
       RENGLON = RENGLON + 1
       FC.Col = 0
       FC.Row = RENGLON
       FC.Text = ""
       FC.Col = 1
       FC.Text = ""
       FC.Col = 2
       FC.Text = ""
       FC.Col = 3
       FC.Text = ""
       FC.Col = 4
       FC.Text = ""
    Loop
RENGLON = 0

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'Esto tiene que funcionar
   cl.Open "SELECT * FROM CATMOVS ORDER BY CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
   cl.MoveFirst
    Do Until cl.EOF = True
            RENGLON = RENGLON + 1
            FC.Col = 0
            FC.Row = RENGLON
            FC.Text = RENGLON
            FC.Col = 1
            FC.Text = cl.Fields("CVEMOV")
            FC.Col = 2
            FC.Text = cl.Fields("TIPO")
            FC.Col = 3
            FC.Text = cl.Fields("APREPAC")
            FC.Col = 4
            FC.Text = cl.Fields("DESCRIP")
        cl.MoveNext
    Loop

cl.Close


End Sub

Private Sub CmdGrabaReg_Click()
       flgsocio = 0

BUSCA_SOCIO
'IntRespuesta = MsgBox("BUSCA Socio: " & n_socio & "-" & "....", 0)
If TxtAPrePac = "P" And SaldoPres < 0.01 And TxtSocio <> "100" Then
    IntRespuesta = MsgBox("Socio: " & n_socio & "-" & "PAGO PRESTAMO NO PROCEDE....", 0)
    Exit Sub
End If

'Exit Sub
'Move the mouse cursor to the point (100,200) on the screen
Dim retval As Long ' return value


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
    .Source = "SELECT * FROM SICMOV"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CursorLocation = adUseServer
    .Open
End With

If TxtImporte > "" Then
    IntRespuesta = MsgBox("Grabar Registro Importe=$ " & TxtImporte, 1)
 If (IntRespuesta = 1) Then
  Do Until cs.EOF = True
    With cs
        .AddNew
        cs.Fields("NUMREG") = cs.Fields("Id")
        s_numreg = cs.Fields("Id")
        cs.Fields("GRUPO") = PrvGrupo
        cs.Fields("SOCIO") = TxtSocio
            n_socio = TxtSocio
        cs.Fields("IMPORTE") = TxtImporte
        cs.Fields("FECHA") = TxtFecha
        Fecha = TxtFecha
        cs.Fields("CVEMOV") = TxtCveMov
        n_cvemov = TxtCveMov
        cs.Fields("TIPO") = TxtTipo
        cs.Fields("APREPAC") = TxtAPrePac
        cs.Fields("DESCRIP") = TxtDescrip
        UBANCO = TxtCtaBco
        UBANCO = UCase(UBANCO)
        TxtCtaBco = UBANCO
        cs.Fields("CTABCO") = UBANCO
        cs.Fields("REFERENC") = TxtReferenc
        If Month(TxtFecha) > 10 Then
            cs.Fields("NUMES") = Month(TxtFecha) - 10
        Else
            cs.Fields("NUMES") = Month(TxtFecha) + 2
        End If
        cs.Fields("TASA") = PrvTasa
        cs.Update
        Exit Do
    End With
    Loop
    If TxtAPrePac = "P" Or TxtAPrePac = "C" Then
        Actualiza_SALDO
        GrabaDMOVPR
    Else
        Actualiza_SALDO
        GrabaDMOVIN
    End If
    If dBanco <> "F3D4M4C" Then
        GrabaDBANCO
    End If
    
    TxtUltImporte = TxtImporte
    TxtUltSocio = n_socio
    TxtImporte = ""
    TxtSocio = n_socio
    TxtCveMov = n_cvemov
    TxtFecha.Text = "  /  /    "
    cs.Close
 Else
    IntRespuesta = MsgBox("El REGISTRO NO FUE GRABADO", 0)
 End If
Else
    IntRespuesta = MsgBox("El Importe no es CORRECTO", 0)
End If

TxtFecha.Text = Fecha
CmdGrabaReg.TabIndex = 0
TxtImporte.SetFocus
TxtImporte.SelStart = Val(TxtImporte.Text)
End Sub
Sub GrabaDBANCO()
Dim cn As ADODB.Connection
Dim cs As ADODB.Recordset
Dim strPath As String
   
'Update the following path to point to the sample
'Northwind.mdb database on your computer.

strPath = "C:\" & Carpeta & "\" & dBanco & ".mdb"
'IntRespuesta = MsgBox(strPath, 0)

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
If TxtImporte > "" Then
'Do Until cs.EOF = True
    With cs
        .AddNew
        cs.Fields("NUMREG") = s_numreg
        cs.Fields("GRUPO") = PrvGrupo
        cs.Fields("SOCIO") = n_socio
        cs.Fields("IMPORTE") = TxtImporte
        cs.Fields("FECHA") = TxtFecha
        cs.Fields("CVEMOV") = TxtCveMov
        cs.Fields("TIPO") = TxtTipo
        cs.Fields("APREPAC") = TxtAPrePac
        cs.Fields("DESCRIP") = TxtDescrip
        cs.Fields("CTABCO") = TxtCtaBco
        cs.Fields("REFERENC") = TxtReferenc
        If Month(TxtFecha) > 10 Then
            cs.Fields("NUMES") = Month(TxtFecha) - 10
        Else
            cs.Fields("NUMES") = Month(TxtFecha) + 2
        End If
        cs.Fields("TASA") = PrvTasa

        cs.Update
        'Exit Do
    End With
    'Loop
Else
    IntRespuesta = MsgBox("El Importe no es CORRECTO", 0)
End If
End Sub
Sub BUSCA_SOCIO()
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."

   cl.MoveFirst
    'IntRespuesta = MsgBox(PubSocio, 0)

    Do Until cl.EOF = True
        'IntRespuesta = MsgBox("Socio: " & TxtSocio & "-" & "TODOS....", 0)
        If TxtSocio < "01" Then
            Exit Do
        End If
        If cl.Fields("SOCIO") = TxtSocio Then
                      'IntRespuesta = MsgBox("Socio: " & n_socio & "-" & "....", 0)
            SaldoPres = cl.Fields("SALDOPRES")
            TxtNombre = cl.Fields("NOMBRE")
           'If TxtAPrePac = "P" And cl.Fields("SALDOPRES") < 0.01 Then
           '   IntRespuesta = MsgBox("Socio: " & n_socio & "-" & "PAGO PRESTAMO NO PROCEDE....", 0)
           'End If
           Exit Do
       End If
       cl.MoveNext
    Loop
'cl.Close
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
        If cl.Fields("SOCIO") = TxtSocio Then
            PrvGrupo = cl.Fields("GRUPO")
            SaldoPres = cl.Fields("SALDOPRES")
            Saldo = cl.Fields("SALDO")
            If TxtAPrePac = "P" Then
                cl.Fields("SALDOPRES") = SaldoPres - TxtImporte
                cl.Fields("PAGOS") = cl.Fields("PAGOS") + TxtImporte
            End If
            If TxtAPrePac = "C" Then
                cl.Fields("SALDOPRES") = SaldoPres + TxtImporte
                cl.Fields("PRESTAMOS") = cl.Fields("PRESTAMOS") + TxtImporte
                If cl.Fields("TIPO") = "4" Then
                    cl.Fields("PAGOMIN") = cl.Fields("SALDOPRES") * 0.05
                End If
            End If
            If TxtAPrePac = "A" Then
                cl.Fields("SALDO") = Saldo + TxtImporte
                cl.Fields("APORTA") = cl.Fields("APORTA") + TxtImporte
            End If
            If TxtAPrePac = "R" Then
                cl.Fields("SALDO") = Saldo - TxtImporte
                cl.Fields("RETIROS") = cl.Fields("RETIROS") + TxtImporte
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
            
            If TxtAPrePac = "P" Then
                cl.Fields("APORTA") = cl.Fields("APORTA") + TxtImporte
                cl.Fields("SALDO") = cl.Fields("SALDO") + TxtImporte
            End If
            If TxtAPrePac = "C" Then
                cl.Fields("RETIROS") = cl.Fields("RETIROS") + TxtImporte
                cl.Fields("SALDO") = cl.Fields("SALDO") - TxtImporte
            End If
            If TxtAPrePac = "A" Then
                cl.Fields("APORTA") = cl.Fields("APORTA") + TxtImporte
                cl.Fields("SALDO") = cl.Fields("SALDO") + TxtImporte
            End If
            If TxtAPrePac = "R" Then
                cl.Fields("RETIROS") = cl.Fields("RETIROS") + TxtImporte
                cl.Fields("SALDO") = cl.Fields("SALDO") - TxtImporte
            End If
            'IntRespuesta = MsgBox("SALDO 99=$" & cl.Fields("SALDO"), 0)

            cl.Update
            TxtSaldoCaja = Format(cl.Fields("SALDO"), "Currency")
            Exit Do
    End If
    cl.MoveNext
        
    Loop
    TxtSocio = ""
End Sub

Sub COLOCATITULOSFS()
FS.Row = 0
FS.Col = 0
FS.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
FS.Col = 1
FS.Text = "SOCIO"
FS.Col = 2
FS.Text = "NOMBRE"
FS.Col = 3
FS.Text = "SALDO"
FS.Col = 4
FS.Text = "PRESTAMO"
FS.Col = 5
FS.Text = "PAGO-MIN"

FS.ColWidth(0) = 450    'AJUSTO EL ANCHO DE LA COLUMNA

FS.ColWidth(1) = 600    'AJUSTO EL ANCHO DE LA COLUMNA
FS.ColWidth(2) = 2000    'AJUSTO EL ANCHO DE LA COLUMNA2

FS.ColWidth(3) = 1200    'AJUSTO EL ANCHO DE LA COLUMNA2
FS.ColWidth(4) = 1200    'AJUSTO EL ANCHO DE LA COLUMNA2

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
If TxtImporte > "" Then
Do Until cs.EOF = True
    With cs
        .AddNew
        cs.Fields("NUMREG") = s_numreg
        cs.Fields("GRUPO") = PrvGrupo
        cs.Fields("SOCIO") = n_socio
        cs.Fields("IMPORTE") = TxtImporte
        cs.Fields("FECHA") = TxtFecha
        cs.Fields("CVEMOV") = TxtCveMov
        cs.Fields("TIPO") = TxtTipo
        cs.Fields("APREPAC") = TxtAPrePac
        cs.Fields("DESCRIP") = TxtDescrip
        cs.Fields("CTABCO") = TxtCtaBco
        cs.Fields("REFERENC") = TxtReferenc
        If Month(TxtFecha) > 10 Then
            cs.Fields("NUMES") = Month(TxtFecha) - 10
        Else
            cs.Fields("NUMES") = Month(TxtFecha) + 2
        End If
        cs.Fields("TASA") = PrvTasa

        cs.Update
        Exit Do
    End With
    Loop
Else
    IntRespuesta = MsgBox("El Importe no es CORRECTO", 0)
End If
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
If TxtImporte > "" Then
Do Until cs.EOF = True
    With cs
        .AddNew
        cs.Fields("NUMREG") = s_numreg
        cs.Fields("GRUPO") = PrvGrupo
        cs.Fields("SOCIO") = n_socio
        cs.Fields("IMPORTE") = TxtImporte
        cs.Fields("FECHA") = TxtFecha
        cs.Fields("CVEMOV") = TxtCveMov
        cs.Fields("TIPO") = TxtTipo
        cs.Fields("APREPAC") = TxtAPrePac
        cs.Fields("DESCRIP") = TxtDescrip
        cs.Fields("CTABCO") = TxtCtaBco
        cs.Fields("REFERENC") = TxtReferenc
        If Month(TxtFecha) > 10 Then
            cs.Fields("NUMES") = Month(TxtFecha) - 10
        Else
            cs.Fields("NUMES") = Month(TxtFecha) + 2
        End If
        cs.Fields("TASA") = PrvTasa

        cs.Update
        Exit Do
    End With
    Loop
Else
    IntRespuesta = MsgBox("El Importe no es CORRECTO", 0)
End If
End Sub

Sub BorraCeldasenFS()
    Do Until RENGLON = 199
       RENGLON = RENGLON + 1
       FS.Col = 0
       FS.Row = RENGLON
       FS.Text = ""
       FS.Col = 1
       FS.Text = ""
       FS.Col = 2
       FS.Text = ""
       FS.Col = 3
       FS.Text = ""
       FS.Col = 4
       FS.Text = ""
    Loop
RENGLON = 0
End Sub

Private Sub CmdSocio_Click()
COLOCATITULOSFS
BorraCeldasenFS
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'ESTA LINEA SOBRA porque ya has especificado el origen en la SELECT:
    'cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"

   'Esto tiene que funcionar
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
   cl.MoveFirst
    LongNom = 1
    LongNom = Len(TxtNomMS)
    Do Until cl.EOF = True
        varnombre = Left(cl.Fields("NOMBRE"), LongNom)
        UNOMBRE = TxtNomMS
        UNOMBRE = UCase(UNOMBRE)

        If varnombre = UNOMBRE Then
            RENGLON = RENGLON + 1
            FS.Col = 0
            FS.Row = RENGLON
            FS.Text = RENGLON
            FS.Col = 1
            FS.Text = cl.Fields("SOCIO")
            FS.Col = 2
            FS.Text = cl.Fields("NOMBRE")
            FS.Col = 3
            FS.Text = Format(cl.Fields("SALDO"), "Currency")
            FS.Col = 4
            FS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
        End If
        cl.MoveNext
    Loop

cl.Close
'ValorFlexGrid

End Sub

Private Sub CmdUltima_Click()
             CtaBco = ""
   
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.SICMOV
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.SOCIOS

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM SICMOV ORDER BY Id", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    'numes = 11
    Do Until cd.EOF = True
        Numovs = Numovs + 1
        prvImporte = cd.Fields("IMPORTE")
        PrvSocio = cd.Fields("SOCIO")
        Fecha = cd.Fields("FECHA")
        CveMov = cd.Fields("CVEMOV")
        Tipo = cd.Fields("TIPO")
        Aprepac = cd.Fields("APREPAC")
        If cd.Fields("CTABCO") <> "" Then
            CtaBco = cd.Fields("CTABCO")
        Else
            CtaBco = "   "
        End If
        If cd.Fields("REFERENC") <> "" Then
            Referenc = cd.Fields("REFERENC")
        Else
            Referenc = "          "
        End If
        If cd.Fields("TASA") <> "" Then
            TxtTasa = cd.Fields("TASA")
        Else
            TxtTasa = 1
        End If
        cd.MoveNext
        Loop
        'IntRespuesta = MsgBox("Numovs = " & Numovs, 0)
    TxtImporte = ""
    TxtUltImporte = prvImporte
    TxtSocio = PrvSocio
    TxtFecha = Fecha
    TxtCveMov = CveMov
    TxtTipo = Tipo
    TxtAPrePac = Aprepac
    'If CtaBco > "" Then
        TxtCtaBco = CtaBco
    'End If
    'If Referenc > "" Then
        TxtReferenc = Referenc
    'End If
    TxtUltSocio = PrvSocio
    
    cl.MoveFirst
   
     Do Until cl.EOF = True
        If cl.Fields("SOCIO") = "99" Then
            TxtSaldoCaja = cl.Fields("SALDO")
            TxtBanco = "LLB"
            Exit Do
        End If
        cl.MoveNext
        Loop
cd.Close
TxtSocioEnter
TxtCveMov = CveMov

TxtCveMovEnter

End Sub

Private Sub FS_Click()
    FS.Col = 1
    TxtSocio = FS.Text
    FS.Col = 2
    'IntRespuesta = MsgBox("Nuevo" & FS.Row & FS.Text, 0)
    FS.Row = 190
    TxtSocioEnter
    TxtSocio.SetFocus
    TxtSocio.SelStart = Val(TxtSocio.Text)
    TxtNomFs = ""
End Sub

Private Sub Label9_Click()
    IntRespuesta = MsgBox("Registra el Importe del Depósito", 0)

End Sub


Private Sub TxtCtaBco_KeyPress(KeyAscii As Integer)
     If TxtCtaBco.Tag = MODE_OVERTYPE And TxtCtaBco.SelLength = 0 Then
        TxtCtaBco.SelLength = 1
    End If
    If (KeyAscii >= 97) And (KeyAscii <= 122) Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       'TxtCveMovEnter
       SendKeys "{tab}"
    End If

End Sub
Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    If TxtDescrip.Tag = MODE_OVERTYPE And TxtDescrip.SelLength = 0 Then
        TxtDescrip.SelLength = 1
    End If
    If (KeyAscii >= 97) And (KeyAscii <= 122) Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       'TxtCveMovEnter
       SendKeys "{tab}"
    End If

End Sub
Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtFecha.Tag = MODE_OVERTYPE And TxtFecha.SelLength = 0 Then
        TxtFecha.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtFechaEnter
       SendKeys "{tab}"
       TxtCveMovEnter
    End If

End Sub
Private Sub TxtFechaEnter()
    If TxtFecha.Tag = MODE_OVERTYPE And TxtFecha.SelLength = 0 Then
        TxtFecha.SelLength = 1
    End If

End Sub

Private Sub TxtNomFs_Change()
COLOCATITULOSFS
BorraCeldasenFS
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'ESTA LINEA SOBRA porque ya has especificado el origen en la SELECT:
    'cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"

   'Esto tiene que funcionar
   cl.Open "SELECT * FROM SOCIOS ORDER BY NOMBRE", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
   cl.MoveFirst
    LongNom = 1
    LongNom = Len(TxtNomFs)
    Do Until cl.EOF = True
        varnombre = Left(cl.Fields("NOMBRE"), LongNom)
        UNOMBRE = TxtNomFs
        UNOMBRE = UCase(UNOMBRE)

        If varnombre = UNOMBRE Then
            RENGLON = RENGLON + 1
            FS.Col = 0
            FS.Row = RENGLON
            FS.Text = RENGLON
            FS.Col = 1
            FS.Text = cl.Fields("SOCIO")
            FS.Col = 2
            FS.Text = cl.Fields("NOMBRE")
            FS.Col = 3
            FS.Text = Format(cl.Fields("SALDO"), "Currency")
            FS.Col = 4
            FS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
            FS.Col = 5
            FS.Text = Format(cl.Fields("PAGOMIN"), "Currency")

        End If
        cl.MoveNext
    Loop

cl.Close


End Sub


Private Sub TxtReferenc_KeyPress(KeyAscii As Integer)
    If TxtReferenc.Tag = MODE_OVERTYPE And TxtReferenc.SelLength = 0 Then
        TxtReferenc.SelLength = 1
    End If
    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       'TxtCveMovEnter
       SendKeys "{tab}"
    End If

End Sub
Private Sub TxtSocio_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtSocio.Tag = MODE_OVERTYPE And TxtSocio.SelLength = 0 Then
        TxtSocio.SelLength = 1
    End If
    TxtSocioEnter

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtSocioEnter
       SendKeys "{tab}"
    End If
    If flgsocio = 0 Then
        If KeyAscii <> 13 Then
            flgsocio = 1
            TxtSocio = ""
        End If
    End If
End Sub

Private Sub TxtSocio_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        TxtNomFs.SetFocus
        TxtNomFs.SelStart = Val(TxtNomFs.Text)
    End If
    If KeyCode = 27 Then
        Unload Me
    End If
        'TxtImporteEnter
        'SendKeys "{tab}"
        'TxtSocioEnter
        'CmdGrabaReg.TabIndex = 8
    'End If
End Sub
Private Sub TxtSocioEnter()

Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."

   cl.MoveFirst
    'IntRespuesta = MsgBox(PubSocio, 0)

    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = TxtSocio Then
          TxtSocioFS = TxtSocio
          TxtNombre = cl.Fields("NOMBRE")
          PrvGrupo = cl.Fields("GRUPO")
          PrvTasa = cl.Fields("TASAPRES")
          TxtTasa = PrvTasa
          If cl.Fields("SALDOPRES") > 0 Then
             TxtCveMov = "50"
          Else
             TxtCveMov = "10"
          End If
          If Right(PrvEnteros, 1) = 0 Then
            TxtCveMov = "10"
          End If
          '       IntRespuesta = MsgBox(Right(PrvEnteros, 1) & TxtCveMov, 0)

          Exit Do
       Else
            TxtNombre = "No existe Nombre de este Socio"

       End If
       cl.MoveNext
    Loop
            If TxtNombre = "No existe Nombre de este Socio" Then
                'IntRespuesta = MsgBox(TxtSocio & "-" & Clave, 0)
                
                TxtSocio = Right(TxtSocio, 2)
                BUSCA_SOCIO
            End If

cl.Close

End Sub
Private Sub FC_Click()
    FC.Col = 1
    TxtCveMov = FC.Text
    TxtCveMovEnter

End Sub
Private Sub TxtCveMov_KeyPress(KeyAscii As Integer)
    If TxtCveMov.Tag = MODE_OVERTYPE And TxtCveMov.SelLength = 0 Then
        TxtCveMov.SelLength = 1
    End If
    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtCveMovEnter
       SendKeys "{tab}"
    End If
    TxtCveMovEnter

End Sub


Private Sub TxtCveMovEnter()

Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "CATMOVS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."

   cl.MoveFirst
    'IntRespuesta = MsgBox(PubSocio, 0)

    Do Until cl.EOF = True
        If cl.Fields("CVEMOV") = TxtCveMov Then
            TxtDescrip = cl.Fields("DESCRIP")
            TxtTipo = cl.Fields("TIPO")
            TxtAPrePac = cl.Fields("APREPAC")
            If TxtTipo = "T" Then
               TxtCtaBco = ""
            End If
          Exit Do
       Else
          TxtDescrip = "No existe esta Clave"
       End If
       cl.MoveNext
    Loop
cl.Close

End Sub
Private Sub TxtImporte_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        TxtNomFs.SetFocus
        TxtNomFs.SelStart = Val(TxtNomFs.Text)
    End If
    If KeyCode = 27 Then
        Unload Me
    End If
        'TxtImporteEnter
        'SendKeys "{tab}"
        'TxtSocioEnter
        'CmdGrabaReg.TabIndex = 8
    'End If
End Sub


Private Sub TxtImporte_KeyDown(KeyCode As Integer, Sjift As Integer)
    If KeyCode = 9 Or KeyAscii = 9 Then
        TxtImporteEnter
        'SendKeys "{tab}"
        TxtSocioEnter
        CmdGrabaReg.TabIndex = 8
    End If
End Sub


Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then 'vbKeyTab
        IntRespuesta = MsgBox("KeyAscii=" & KeyAscii & "-" & prvImporte, 0)
    End If
    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtImporteEnter


       SendKeys "{tab}"
       TxtSocioEnter
       CmdGrabaReg.TabIndex = 8
    End If

End Sub
Private Sub TxtImporteEnter()

    prvImporte = TxtImporte
    
    If TxtImporte > "" Then
        Clave = Right(prvImporte, 2)

        PrvEnteros = Right(prvImporte, 4)
        PrvEnteros = Left(PrvEnteros, 1)
    End If
    
    If Right(PrvEnteros, 1) = 1 Then
       Clave = Right(PrvEnteros, 1) & Right(prvImporte, 2)
       TxtCveMov = "50"
    End If
    If Right(PrvEnteros, 1) = 0 Then
       Clave = Right(prvImporte, 2)
       TxtCveMov = "10"
    End If
    If Right(PrvEnteros, 1) = 2 Then
        Clave = Right(PrvEnteros, 1) & Right(prvImporte, 2)
        TxtCveMov = "10"
    End If
    If Right(PrvEnteros, 1) = 3 Then
       Clave = Right(PrvEnteros, 1) - 1 & Right(prvImporte, 2)
       TxtCveMov = "50"
    Else
        TxtCveMov = "10"
    End If
    TxtSocio = Clave
    ESPERA_SOCIO
    If TxtNombre = "No existe Nombre de este Socio" Then

        Clave = (Right(PrvEnteros, 1) - 1) & Right(prvImporte, 2)
        If Left(Clave, 1) = "0" Then
            Clave = Right(prvImporte, 2)
            TxtCveMov = "10"
        End If
       ESPERA_SOCIO

    End If

End Sub
Private Sub ESPERA_SOCIO()

    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."

   cl.MoveFirst
    'IntRespuesta = MsgBox(PubSocio, 0)

    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = TxtSocio Then
          TxtNombre = cl.Fields("NOMBRE")
          PrvGrupo = cl.Fields("GRUPO")
          PrvTasa = cl.Fields("TASAPRES")
          TxtTasa = PrvTasa

          If Right(PrvEnteros, 1) = 1 Then
             If TxtCveMov = "50" Then
                If cl.Fields("SALDOPRES") < 0.01 Then
                    'IntRespuesta = MsgBox("Socio.-" & TxtSocio & "PAGO DE PRESTAMO NO PROCEDE", 0) '
                    TxtSocioFS = Right(PrvEnteros, 1) & Right(prvImporte, 2)
                    'Clave = Right(PrvEnteros, 1) & Right(prvImporte, 2)
                    TxtSocio = TxtSocioFS
                    Exit Sub
                End If
             End If
          End If

          If Right(PrvEnteros, 1) = 2 And cl.Fields("SALDOPRES") < 0.01 Then
             TxtCveMov = "10"
          End If

          If Right(PrvEnteros, 1) = 0 Then
             TxtCveMov = "10"
          End If

          Exit Do
       Else

            TxtNombre = "No existe Nombre de este Socio"
            If TxtNombre = "No existe Nombre de este Socio" Then

                'Clave = (Right(PrvEnteros, 1) - 1) & Right(prvImporte, 2)
                If Left(TxtSocio, 1) = "1" Then
                    Clave = Right(TxtSocio, 2)
                    TxtCveMov = "10"
                End If
            End If
            'ESPERA_SOCIO

       End If
       cl.MoveNext
    Loop

End Sub

Private Sub TxtSocioFS_Change()
COLOCATITULOSFS
BorraCeldasenFS
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'ESTA LINEA SOBRA porque ya has especificado el origen en la SELECT:
    'cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"

   'Esto tiene que funcionar
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
   cl.MoveFirst
    LongNom = 1
    LongNom = Len(TxtNomMS)
    Do Until cl.EOF = True
        varnombre = Left(cl.Fields("NOMBRE"), LongNom)
        UNOMBRE = TxtNomMS
        UNOMBRE = UCase(UNOMBRE)
    If TxtSocioFS <= cl.Fields("SOCIO") Then
            RENGLON = RENGLON + 1
            FS.Col = 0
            FS.Row = RENGLON
            FS.Text = RENGLON
            FS.Col = 1
            FS.Text = cl.Fields("SOCIO")
            FS.Col = 2
            FS.Text = cl.Fields("NOMBRE")
            FS.Col = 3
            FS.Text = Format(cl.Fields("SALDO"), "Currency")
            FS.Col = 4
            FS.Text = Format(cl.Fields("SALDOPRES"), "Currency")
        End If
        cl.MoveNext
    Loop

cl.Close
'ValorFlexGrid

End Sub

Private Sub TxtUltSocio_Change()
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."

   cl.MoveFirst
    'IntRespuesta = MsgBox(PubSocio, 0)

    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = TxtUltSocio Then
          TxtUltNombre = cl.Fields("NOMBRE")
          Exit Do
       Else
            TxtUltNombre = "No existe Nombre de este Socio"
       End If
       cl.MoveNext
    Loop
cl.Close

End Sub


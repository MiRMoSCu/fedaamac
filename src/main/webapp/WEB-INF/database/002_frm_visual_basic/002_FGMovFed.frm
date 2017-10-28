VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FG 
   Caption         =   "Movimientos de Fedamac"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdDepEfectivo 
      Caption         =   "Depósitos de Recuperación"
      Height          =   555
      Left            =   9600
      TabIndex        =   29
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton CmdListAgenda 
      Caption         =   "Lista Agenda"
      Height          =   375
      Left            =   11880
      TabIndex        =   28
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox TxtTipomov 
      Height          =   375
      Left            =   8640
      TabIndex        =   25
      Top             =   6360
      Width           =   615
   End
   Begin VB.TextBox TxtDate 
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
      Left            =   11640
      TabIndex        =   24
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtModFec 
      Height          =   285
      Left            =   1320
      TabIndex        =   22
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CmdModReg 
      Caption         =   "Modifica Datos del Registro"
      Height          =   495
      Left            =   10320
      TabIndex        =   21
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox TxtModBco 
      Height          =   285
      Left            =   6840
      TabIndex        =   20
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox TxtModDescrip 
      Height          =   285
      Left            =   3360
      TabIndex        =   19
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox TxtModRef 
      Height          =   285
      Left            =   5640
      TabIndex        =   18
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox TxtModImp 
      Height          =   285
      Left            =   8040
      TabIndex        =   17
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   7680
      TabIndex        =   15
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton CmdEliminaMov 
      Caption         =   "Elimina Movimiento Seleccionado"
      Height          =   735
      Left            =   3840
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton CmdListMov 
      Caption         =   "Lista Movimientos Capturados"
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton CmdCatMovs 
      Caption         =   "Consulta Catálogo de Movimientos"
      Height          =   855
      Left            =   11880
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton CmdBuscaImp 
      Caption         =   "Busca Importe"
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox TxtImport 
      BeginProperty DataFormat 
         Type            =   5
         Format          =   "0.00"
         HaveTrueFalseNull=   1
         TrueValue       =   "True"
         FalseValue      =   "False"
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   7
      EndProperty
      Height          =   285
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   10
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton CmdTrasp 
      Caption         =   "Movimientos de Traspaso"
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox TxtBco 
      Height          =   285
      Left            =   6360
      TabIndex        =   6
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton CmdMovPres 
      Caption         =   "Consulta Movimientos de Préstamos"
      Height          =   735
      Left            =   11880
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton CmdMovFG 
      Caption         =   "Consulta Movimientos de Inversión"
      Height          =   735
      Left            =   11880
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox TxtMovNom 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox TxtMovSoc 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   5175
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   2000
      Cols            =   12
   End
   Begin VB.Label LblEmpresa 
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Tipo de Movimiento"
      Height          =   375
      Left            =   8640
      TabIndex        =   26
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha de Proceso:"
      Height          =   255
      Left            =   10200
      TabIndex        =   23
      Top             =   240
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   9240
      X2              =   10320
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      Caption         =   "Selecciona Nombre"
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CONSULTA MOVIMIENTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Selecciona Banco"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "SOCIO"
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "FG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Carpeta, s_numreg As String
Private sdoActual As String
Private Importe As Single
Private UltImporte, prvImporte As Single
Private UltSocio, PrvSocio As String
Private PrvFecha As String
Private PrvCveMov As String
Private PrvReferenc As String
Private PrvAPrePac, PrvDescrip As String
Private Numovs, RenBorra, Flag_Mod, FlgBco, FlgTrasp, Flg_movs As Single
Private TotEfectivo, TotDoctos, PrvNumes, BcoEfectivo, BcoDoctos As Single
Private MesEfectivo, MesDoctos As Single
Private Fecha As Date




    'IntRespuesta = MsgBox(prvImporte & "-" & cd.Fields("IMPORTE"), 1)


Sub COLOCATITULOSENFG()
Numovs = 2000
FG.Row = 0
FG.Col = 0
FG.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
FG.Col = 1
FG.Text = "SOCIO"
FG.Col = 2
FG.Text = "FECHA"
FG.Col = 3
FG.Text = "CVE"
FG.Col = 4
FG.Text = "T"
FG.Col = 5
FG.Text = "C"
FG.Col = 6
FG.Text = "DESCRIPCION"
FG.Col = 7
FG.Text = "REFERENCIA"
FG.Col = 8
FG.Text = "BANCO"
FG.Col = 9
FG.Text = "DEPOSITOS"
FG.Col = 10
FG.Text = "RETIROS"
FG.Col = 11
FG.Text = "SALDO"
                    'AJUSTO EL ANCHO DE LAS COLUMNAS
FG.ColWidth(0) = 450
               
FG.ColWidth(1) = 600
FG.ColWidth(2) = 1000

FG.ColWidth(3) = 500
FG.ColWidth(4) = 200
FG.ColWidth(5) = 200

FG.ColWidth(6) = 2700
FG.ColWidth(7) = 1150

FG.ColWidth(8) = 650
FG.ColWidth(9) = 1100
FG.ColWidth(10) = 1100

FG.ColWidth(11) = 1300

sdoActual = 0
End Sub
Sub BorraCeldasenFG()
    FG.Rows = 2000
    RENGLON = RenBorra
    'MsgBox ("RenBorra=" & RenBorra & " FG.Rows=" & FG.Rows)
    Do Until RENGLON = Numovs - 1
       RENGLON = RENGLON + 1
       FG.Col = 0
       FG.Row = RENGLON
       FG.Text = ""
       FG.Col = 1
       FG.Text = ""
       FG.Col = 2
       FG.Text = ""
       FG.Col = 3
       FG.Text = ""
       FG.Col = 4
       FG.Text = ""
       FG.Col = 5
       FG.Text = ""
       FG.Col = 6
       FG.Text = ""
       FG.Col = 7
       FG.Text = ""
       FG.Col = 8
       FG.Text = ""
       FG.Col = 9
       FG.Text = ""
       FG.Col = 10
       FG.Text = ""
       FG.Col = 11
       FG.Text = ""

    Loop

End Sub



Private Sub CmdBuscaImp_Click()
       prvImporte = TxtImport
       RenBorra = 2000 - RenBorra
       BuscaImp

End Sub

Private Sub CmdDepEfectivo_Click()
COLOCATITULOSENFG
'MsgBox ("RenBorra=" & RenBorra)
RenBorra = 1
FG.Rows = 2000
BorraCeldasenFG
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM SICMOV ORDER BY NUMES,CTABCO,FECHA,REFERENC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    Numovs = 0
    Do Until cd.EOF = True
        If cd.Fields("CVEMOV") = "10" Or cd.Fields("CVEMOV") = "50" _
            Or cd.Fields("CVEMOV") = "11" Or cd.Fields("CVEMOV") = "51" _
            Or cd.Fields("CVEMOV") = "12" Or cd.Fields("CVEMOV") = "52" Then
            
            Numovs = Numovs + 1
            If cd.Fields("CTABCO") <> PrvCtaBco Then
                Numovs = Numovs + 3.5
                PrvCtaBco = cd.Fields("CTABCO")
            End If
        End If
        cd.MoveNext
    Loop
    PrvNumes = 0
    
    FG.Row = 1
    'FG.TopRow = 1
    FG.TopRow = Numovs
    cd.MoveFirst
    
    
    
    Do Until cd.EOF = True
       If cd.Fields("CTABCO") > "AAA" Then
        If cd.Fields("CTABCO") <> PrvCtaBco Then
            PrvCtaBco = cd.Fields("CTABCO")
            RENGLON = RENGLON + 1
            FG.Row = RENGLON
            FG.Col = 6
            FG.Text = "DEPOSITOS EN EFECTIVO"
            FG.Col = 10
            FG.Text = Format(BcoEfectivo, "Currency")
            RENGLON = RENGLON + 1
            FG.Row = RENGLON
            FG.Col = 6
            FG.Text = "DEPOSITOS CON DOCTOS"
            FG.Col = 10
            FG.Text = Format(BcoDoctos, "Currency")
            RENGLON = RENGLON + 1
            MesEfectivo = MesEfectivo + BcoEfectivo
            MesDoctos = MesDoctos + BcoDoctos
            BcoEfectivo = 0
            BcoDoctos = 0
        End If
        If cd.Fields("NUMES") <> PrvNumes Then
            PrvNumes = cd.Fields("NUMES")
            RENGLON = RENGLON + 1
            FG.Row = RENGLON
            FG.Col = 6
            FG.Text = "TOT DEPOSITOS EN EFECTIVO"
            FG.Col = 10
            FG.Text = Format(MesEfectivo, "Currency")
            RENGLON = RENGLON + 1
            FG.Row = RENGLON
            FG.Col = 6
            FG.Text = "TOT DEPOSITOS CON DOCTOS"
            FG.Col = 10
            FG.Text = Format(MesDoctos, "Currency")
            RENGLON = RENGLON + 1
            FG.Row = RENGLON
            FG.Col = 6
            FG.Text = "DEPOSITOS TOTALES AL MES"
            FG.Col = 10
            FG.Text = Format(MesDoctos + MesEfectivo, "Currency")
            RENGLON = RENGLON + 1
            MesEfectivo = 0
            MesDoctos = 0
            
        End If
        If cd.Fields("CVEMOV") = "10" Or cd.Fields("CVEMOV") = "50" _
            Or cd.Fields("CVEMOV") = "11" Or cd.Fields("CVEMOV") = "51" _
            Or cd.Fields("CVEMOV") = "12" Or cd.Fields("CVEMOV") = "52" Then
            
            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = cd.Fields("NUMREG")
            FG.Col = 1
            FG.Text = cd.Fields("SOCIO")
            FG.Col = 2
            FG.Text = cd.Fields("FECHA")
            FG.Col = 3
            FG.Text = cd.Fields("CVEMOV")
            FG.Col = 4
            FG.Text = cd.Fields("TIPO")
            FG.Col = 5
            FG.Text = cd.Fields("APREPAC")
            FG.Col = 6
            FG.Text = cd.Fields("DESCRIP")
            FG.Col = 7
            If cd.Fields("REFERENC") > 0 Then
                FG.Text = cd.Fields("REFERENC")
            End If
            FG.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FG.Text = cd.Fields("CTABCO")
            End If
            FG.Col = 9
            FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
        End If
        If cd.Fields("CVEMOV") = "10" Or cd.Fields("CVEMOV") = "50" Then
            BcoEfectivo = BcoEfectivo + cd.Fields("IMPORTE")
        End If
        If cd.Fields("CVEMOV") = "11" Or cd.Fields("CVEMOV") = "51" _
            Or cd.Fields("CVEMOV") = "12" Or cd.Fields("CVEMOV") = "52" Then
            BcoDoctos = BcoDoctos + cd.Fields("IMPORTE")
        End If
       End If
    cd.MoveNext

Loop
RENGLON = RENGLON + 1
FG.Row = RENGLON
FG.Col = 6
FG.Text = "DEPOSITOS EN EFECTIVO"
FG.Col = 10
FG.Text = Format(BcoEfectivo, "Currency")
RENGLON = RENGLON + 1
FG.Row = RENGLON
FG.Col = 6
FG.Text = "DEPOSITOS CON DOCTOS"
FG.Col = 10
FG.Text = Format(BcoDoctos, "Currency")
RENGLON = RENGLON + 1

RENGLON = RENGLON + 1
FG.Row = RENGLON
FG.Col = 6
FG.Text = "TOTAL DEPOSITOS EN EFECTIVO"
FG.Col = 10
FG.Text = Format(MesEfectivo, "Currency")
RENGLON = RENGLON + 1
FG.Row = RENGLON
FG.Col = 6
FG.Text = "TOTAL DEPOSITOS CON DOCTOS"
FG.Col = 10
FG.Text = Format(MesDoctos, "Currency")
RENGLON = RENGLON + 1
FG.Row = RENGLON
FG.Col = 6
FG.Text = "TOTAL RECUPERACION AL MES"
FG.Col = 10
FG.Text = Format(MesDoctos + MesEfectivo, "Currency")
RENGLON = RENGLON + 1

cd.Close
End Sub

Private Sub CmdModReg_Click()
    If s_numreg = "0" Then
        IntRespuesta = MsgBox("ESTE REGISTROS NO SE PUEDE MODIFICAR", 0)
    Else
        IntRespuesta = MsgBox("El Registro Num:" & s_numreg & " Socio:" & PrvSocio & " Importe:" & prvImporte & " Se va a Modificar?", 1)
        If (IntRespuesta = 1) Then
            Flag_Mod = 1
            BuscaRegModificado
            IntRespuesta = MsgBox("EL REGISTRO FUE MODIFICADO:-" & PrvAPrePac & " SOCIO Num:" & PrvSocio & " Importe:" & TxtModImp, 0)

            Actualiza_SOCIOS
        Else
            IntRespuesta = MsgBox("El Registro Num:" & s_numreg & " Socio:" & PrvSocio & " Importe:" & prvImporte & " NO FUE MODIFICADO", 0)
        End If
    End If
If FlgBco = 1 Then
    TxtBcoEnter
    FlgBco = 0
End If
If FlgTrasp = 1 Then
    CmdTrasp_Click
    FlgTrasp = 0
End If
Flg_movs = 1
CmdListMov_Click
Flg_movs = 0
End Sub
Private Sub BuscaRegModificado()

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

Do Until cs.EOF = True
    If s_numreg = cs.Fields("NUMREG") Then
        
        'IntRespuesta = MsgBox("SICMOV El Registro Num:" & s_numreg & " Importe:" & prvImporte & " Se va a Modificar?", 1)

        cs.Fields("FECHA") = TxtModFec
        cs.Fields("REFERENC") = TxtModRef
        cs.Fields("DESCRIP") = TxtModDescrip
        cs.Fields("CTABCO") = TxtModBco
        cs.Fields("IMPORTE") = TxtModImp
        cs.Update
        Exit Do
    End If

    cs.MoveNext
    Loop
    If PrvAPrePac = "A" Or PrvAPrePac = "R" Then
        ACTUALIZA_DMOVIN
    Else
        ACTUALIZA_DMOVPR
    End If
cs.Close
'CmdListMov_Click
End Sub

Sub ACTUALIZA_DMOVIN()

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
    If s_numreg = cs.Fields("NUMREG") Then
        'IntRespuesta = MsgBox("DMOVIN El Registro Num:" & s_numreg & " Importe:" & prvImporte & " Se va a Modificar?", 1)

        cs.Fields("FECHA") = TxtModFec
        cs.Fields("REFERENC") = TxtModRef
        cs.Fields("DESCRIP") = TxtModDescrip
        cs.Fields("CTABCO") = TxtModBco
        cs.Fields("IMPORTE") = TxtModImp
        cs.Update
        Exit Do
    End If

    cs.MoveNext
    Loop
End Sub
Sub ACTUALIZA_DMOVPR()

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
    If s_numreg = cs.Fields("NUMREG") Then
        'IntRespuesta = MsgBox("DMOVPR El Registro Num:" & s_numreg & " Importe:" & prvImporte & " Se va a Modificar?", 1)

        cs.Fields("FECHA") = TxtModFec
        cs.Fields("REFERENC") = TxtModRef
        cs.Fields("DESCRIP") = TxtModDescrip
        cs.Fields("CTABCO") = TxtModBco
        cs.Fields("IMPORTE") = TxtModImp
        cs.Update
        Exit Do
    End If

    cs.MoveNext
    Loop
End Sub

Private Sub Command1_Click()
    Dim Word As Object
    Set Word = CreateObject("Word.Application")

    'Dim Word As New Word.Application

    'AGREGA  DOCUMENTO
    Word.Documents.Add
    
    Dim x As Integer
    For x = 1 To 10
        'AGREGA TEXTO
        'Word.Selection.Font.Color = wdColorRed
        Word.Selection.TypeText "FONDO ECONOMICO DE AYUDA MUTUA, A.C. " & vbCrLf
        Word.Selection.TypeText "        ESTADO DE CUENTA "
        Word.Selection.TypeText "Hola," + "Espero que todo vaya bien." + "Saludos, Línea de Código."
        'AGREGA PARRAFO
        Word.Selection.TypeParagraph
    Next
    
    'SELECCIONA TEXTO
    Word.Selection.WholeStory
    Word.Selection.Font.Size = 14
    
    
    ' VISIBLE
    Word.Visible = True

    Set Word = Nothing

End Sub





Private Sub TxtImport_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       prvImporte = TxtImport
       RenBorra = 2000 - RenBorra
       BuscaImp
    End If

End Sub

'Private Sub TxtImportEnter()
'    Dim x As Single
'    x = 0
'    For x = 0 To 10
'        prvImporte = TxtImport
'        Next x
    
    
     'IntRespuesta = MsgBox(x & "-" & privImporte & "-" & TxtImport, 1)
   
'End Sub



Private Sub BuscaImp()
COLOCATITULOSENFG
FG.Rows = 2000
BorraCeldasenFG

    varImporte = TxtImporte

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM SICMOV ORDER BY IMPORTE,CTABCO,FECHA,REFERENC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    FG.Row = 0
    FG.TopRow = 1
    'IntRespuesta = MsgBox(prvImporte & "-" & cd.Fields("IMPORTE"), 1)
    Do Until cd.EOF = True
        Importe = cd.Fields("IMPORTE")
        If Importe >= prvImporte Then
            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = cd.Fields("NUMREG")
            FG.Col = 1
            FG.Text = cd.Fields("SOCIO")
            FG.Col = 2
            FG.Text = cd.Fields("FECHA")
            FG.Col = 3
            FG.Text = cd.Fields("CVEMOV")
            FG.Col = 4
            FG.Text = cd.Fields("TIPO")
            FG.Col = 5
            FG.Text = cd.Fields("APREPAC")

            FG.Col = 6
            FG.Text = cd.Fields("DESCRIP")
            FG.Col = 7
            FG.Text = ""
            FG.Col = 8
            FG.Text = ""
            FG.Col = 9
            FG.Text = ""
            FG.Col = 10
            FG.Text = ""

            If cd.Fields("REFERENC") > 0 Then
                FG.Text = cd.Fields("REFERENC")
            End If
            FG.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FG.Text = cd.Fields("CTABCO")
            End If
            If (cd.Fields("CVEMOV") < "20" Or cd.Fields("CVEMOV") = "48" Or cd.Fields("CVEMOV") = "69") Or (cd.Fields("CVEMOV") > "49" And cd.Fields("CVEMOV") < "59") Then
                    '      *Abonos
                FG.Col = 9
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual + cd.Fields("IMPORTE")
            Else
                '      *Cargos
                FG.Col = 10
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual - cd.Fields("IMPORTE")
            End If
        
        End If

    cd.MoveNext
Loop
TxtImport = ""
End Sub

Private Sub CmdCatMovs_Click()

COLOCATITULOSENFG
RenBorra = 0
FG.Rows = 2000
FG.TopRow = 1
BorraCeldasenFG

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM CATMOVS ORDER BY CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    numes = 11
    Do Until cd.EOF = True

            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = RENGLON
            FG.Col = 3
            FG.Text = cd.Fields("CVEMOV")
            FG.Col = 4
            FG.Text = cd.Fields("TIPO")
            FG.Col = 5
            FG.Text = cd.Fields("APREPAC")
            FG.Col = 6
            FG.Text = cd.Fields("DESCRIP")

    cd.MoveNext
Loop
cd.Close
End Sub

Private Sub CmdEliminaMov_Click()
CuentaMovs
'MsgBox ("Numovs=" & Numovs)

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

Do Until cs.EOF = True
     If cs.Fields("SOCIO") = PrvSocio Then

        If cs.Fields("FECHA") = PrvFecha Then
            If cs.Fields("CVEMOV") = PrvCveMov Then
                If cs.Fields("IMPORTE") = prvImporte Then
                    'MsgBox ("Numovs=" & Numovs)
                    IntRespuesta = MsgBox(PrvSocio & ".-" & cs.Fields("DESCRIP") & " $" & cs.Fields("IMPORTE") & ".-El Movimiento seleccionado Será borrado", 1)
                    Exit Do
                End If
            End If
        End If
    End If

    cs.MoveNext
Loop
    If Not cs.EOF Then
        If (IntRespuesta = 1) Then
            'BorraDETMOV
            If cs.Fields("APREPAC") = "P" Or cs.Fields("APREPAC") = "C" Then
                BorraDMOVPR
                Actualiza_SOCIOS
            Else
                BorraDMOVIN
                Actualiza_SOCIOS
            End If
            cs.Delete
            IntRespuesta = MsgBox("El Movimiento seleccionado FUE BORRADO...", 0)

        Else
            IntRespuesta = MsgBox("El Registro NO FUE BORRADO", 0)
        End If
    End If

    'MsgBox ("Numovs=" & Numovs)
    Flg_movs = 1
CmdListMov_Click
Flg_movs = 0
'RenBorra = RENGLON
'BorraCeldasenFG
cs.Close

End Sub
Sub BorraDETMOV()

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
    .Source = "SELECT * FROM DETMOV"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CursorLocation = adUseServer
    .Open
End With

Do Until cs.EOF = True
    If cs.Fields("SOCIO") = PrvSocio Then
        If cs.Fields("FECHA") = PrvFecha Then
            If cs.Fields("CVEMOV") = PrvCveMov Then
                If cs.Fields("IMPORTE") = prvImporte Then
                    'IntRespuesta = MsgBox("DETMOV" & cs.Fields("DESCRIP") & " $" & cs.Fields("IMPORTE") & ".-El Movimiento seleccionado Será borrado", 1)
                    cs.Delete
                    Exit Do
                End If
            End If
        End If
    End If

    cs.MoveNext
    Loop
End Sub
Sub BorraDMOVPR()

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
    If cs.Fields("SOCIO") = PrvSocio Then
        If cs.Fields("FECHA") = PrvFecha Then
            If cs.Fields("CVEMOV") = PrvCveMov Then
                If cs.Fields("IMPORTE") = prvImporte Then
                    'IntRespuesta = MsgBox("DMOVPR" & cs.Fields("DESCRIP") & " $" & cs.Fields("IMPORTE") & ".-El Movimiento seleccionado Será borrado", 1)
                    cs.Delete
                    Exit Do
                End If
            End If
        End If
    End If

    cs.MoveNext
    Loop
End Sub
Sub BorraDMOVIN()

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
    If cs.Fields("SOCIO") = PrvSocio Then
        If cs.Fields("FECHA") = PrvFecha Then
            If cs.Fields("CVEMOV") = PrvCveMov Then
                If cs.Fields("IMPORTE") = prvImporte Then
                    'IntRespuesta = MsgBox("DMOVIN" & cs.Fields("DESCRIP") & " $" & cs.Fields("IMPORTE") & ".-El Movimiento seleccionado Será borrado", 1)
                    cs.Delete
                    Exit Do
                End If
            End If
        End If
    End If

    cs.MoveNext
    Loop
End Sub
Private Sub CmdListMov_Click()
COLOCATITULOSENFG
frmMiPrimera.Flg = 0
Label3 = "CONSULTA MOVIMIENTOS"

CuentaMovs


'BorraCeldasenFG


   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM SICMOV ORDER BY Id", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    FG.Rows = Numovs
    If Flg_movs = 0 Then
        FG.TopRow = Numovs - 1
    End If
    
    'MsgBox ("Numovs=" & Numovs)

    'IntRespuesta = MsgBox(privImporte & "-" & cd.Fields("IMPORTE"), 1)
    Do Until cd.EOF = True
        Fecha = cd.Fields("FECHA")
        'If Fecha > TxtDate Then
            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = cd.Fields("NUMREG")
            FG.Col = 1
            FG.Text = cd.Fields("SOCIO")
            FG.Col = 2
            FG.Text = cd.Fields("FECHA")
            FG.Col = 3
            FG.Text = cd.Fields("CVEMOV")
            FG.Col = 4
            FG.Text = cd.Fields("TIPO")
            FG.Col = 5
            FG.Text = cd.Fields("APREPAC")

            FG.Col = 6
            FG.Text = cd.Fields("DESCRIP")
            FG.Col = 7
            If cd.Fields("REFERENC") > 0 Then
                FG.Text = cd.Fields("REFERENC")
            End If
            FG.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FG.Text = cd.Fields("CTABCO")
            End If
            If (cd.Fields("CVEMOV") < "20" Or cd.Fields("CVEMOV") = "48" Or cd.Fields("CVEMOV") = "69") Or (cd.Fields("CVEMOV") > "49" And cd.Fields("CVEMOV") < "59") Then
                    '      *Abonos
                FG.Col = 9
                FG.Text = cd.Fields("IMPORTE")
                sdoActual = sdoActual + cd.Fields("IMPORTE")
                FG.Col = 10
                FG.Text = ""
            Else
                '      *Cargos
                FG.Col = 10
                FG.Text = cd.Fields("IMPORTE")
                sdoActual = sdoActual - cd.Fields("IMPORTE")
                FG.Col = 9
                FG.Text = ""
            End If
    'End If
    cd.MoveNext
Loop
RenBorra = RENGLON
BorraCeldasenFG


End Sub
Private Sub CmdListAgenda_Click()
Numovs = 100
FG.Row = 0
FG.Col = 0
FG.Text = "Numreg"         'SE TRATA DE NUMERAR LOS REGISTROS
FG.Col = 1
FG.Text = "FECHA"
FG.Col = 2
FG.Text = "HORA"
FG.Col = 3
FG.Text = "ASUNTO"
FG.Col = 4
FG.Text = "UBICACION"
FG.Col = 5
FG.Text = "INDICACIONES"
FG.Col = 6
FG.Text = "RECORDATORIO"

                    'AJUSTO EL ANCHO DE LAS COLUMNAS
FG.ColWidth(0) = 450
               
FG.ColWidth(1) = 1000
FG.ColWidth(2) = 600

FG.ColWidth(3) = 3400
FG.ColWidth(4) = 3400
FG.ColWidth(5) = 2400
FG.ColWidth(6) = 1400

'BorraCeldasenFG


   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Agenda

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & "SYSFED" & "\Agenda.mdb"
   
   cd.Open "SELECT * FROM AGENDA ORDER BY FECHA", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    FG.Rows = Numovs

    'IntRespuesta = MsgBox("Numovs = " & Numovs, 0)

    'IntRespuesta = MsgBox(privImporte & "-" & cd.Fields("IMPORTE"), 1)
    Do Until cd.EOF = True
        Fecha = cd.Fields("FECHA")
            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = cd.Fields("Id")
            FG.Col = 1
            FG.Text = cd.Fields("FECHA")
            FG.Col = 2
            FG.Text = cd.Fields("HORA") & ":" & cd.Fields("MINUTOS")
            FG.Col = 3
            FG.Text = cd.Fields("ASUNTO")
            FG.Col = 4
            FG.Text = cd.Fields("UBICA")
            FG.Col = 5
            FG.Text = cd.Fields("INDICA")

            FG.Col = 6
            FG.Text = cd.Fields("FECHAR")
    cd.MoveNext
Loop
RenBorra = RENGLON
BorraCeldasenFG


End Sub

Private Sub CmdMovFG_Click()
COLOCATITULOSENFG
FG.Rows = 2000
'BorraCeldasenFG

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM DMOVIN ORDER BY SOCIO,FECHA,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    numes = 11
    Do Until cd.EOF = True
        If cd.Fields("SOCIO") = TxtMovSoc Then
            'IntRespuesta = MsgBox("SOCIO=" & cd.Fields("CVEMOV") & cd.Fields("IMPORTE"), 0)

            If Month(cd.Fields("FECHA")) <> numes Then
                numes = Month(cd.Fields("FECHA"))
                RENGLON = RENGLON + 1
                FG.Col = 0
                FG.Row = RENGLON
                FG.Text = RENGLON
                FG.Col = 1
                FG.Text = ""
                FG.Col = 2
                FG.Text = ""
                FG.Col = 3
                FG.Text = ""
                FG.Col = 4
                FG.Text = ""
                FG.Col = 5
                FG.Text = ""
                FG.Col = 6
                FG.Text = ""
                FG.Col = 7
                FG.Text = ""
                FG.Col = 8
                FG.Text = ""
                FG.Col = 9
                FG.Text = ""
                FG.Col = 10
                FG.Text = ""
                FG.Col = 11
                FG.Text = ""
            
            End If

            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = cd.Fields("NUMREG")
            FG.Col = 1
            FG.Text = cd.Fields("SOCIO")
            FG.Col = 2
            FG.Text = cd.Fields("FECHA")
            FG.Col = 3
            FG.Text = cd.Fields("CVEMOV")
            FG.Col = 4
            FG.Text = ""
            FG.Col = 5
            FG.Text = ""

            FG.Col = 6
            FG.Text = cd.Fields("DESCRIP")
            FG.Col = 7
            If cd.Fields("REFERENC") > 0 Then
                FG.Text = cd.Fields("REFERENC")
            End If
            FG.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FG.Text = cd.Fields("CTABCO")
            End If
            If cd.Fields("APREPAC") = "A" Then
                '      *Abonos
                FG.Col = 9
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                FG.Col = 10
                FG.Text = ""
                sdoActual = sdoActual + cd.Fields("IMPORTE")
            Else
                '      *Cargos
                FG.Col = 9
                FG.Text = ""
                FG.Col = 10
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual - cd.Fields("IMPORTE")
            End If
            FG.Col = 11
            FG.Text = Format(sdoActual, "Currency")
            
        End If

    cd.MoveNext
    
Loop
RenBorra = RENGLON
BorraCeldasenFG

cd.Close

End Sub




Private Sub CmdMovPres_Click()
COLOCATITULOSENFG
FG.Col = 9
FG.Text = "PAGOS"
FG.Col = 10
FG.Text = "PRESTAMOS"
'BorraCeldasenFG

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM DMOVPR ORDER BY SOCIO,FECHA,APREPAC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    numes = 11
    Do Until cd.EOF = True
        If cd.Fields("SOCIO") = TxtMovSoc Then
           If Month(cd.Fields("FECHA")) <> numes Then
                numes = Month(cd.Fields("FECHA"))
                RENGLON = RENGLON + 1
                FG.Col = 0
                FG.Row = RENGLON

                FG.Text = cd.Fields("NUMREG")
                FG.Col = 1
                FG.Text = ""
                FG.Col = 2
                FG.Text = ""
                FG.Col = 3
                FG.Text = ""
                FG.Col = 4
                FG.Text = ""
                FG.Col = 5
                FG.Text = ""
                FG.Col = 6
                FG.Text = ""
                FG.Col = 7
                FG.Text = ""
                FG.Col = 8
                FG.Text = ""
                FG.Col = 9
                FG.Text = ""
                FG.Col = 10
                FG.Text = ""
                FG.Col = 11
                FG.Text = ""
            End If

            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = cd.Fields("NUMREG")
            FG.Col = 1
            FG.Text = cd.Fields("SOCIO")
            FG.Col = 2
            FG.Text = cd.Fields("FECHA")
            FG.Col = 3
            FG.Text = cd.Fields("CVEMOV")
            FG.Col = 4
            FG.Text = ""
            FG.Col = 5
            FG.Text = ""

            FG.Col = 6
                FG.Text = cd.Fields("DESCRIP")
            FG.Col = 7
                If cd.Fields("REFERENC") > 0 Then
                    FG.Text = cd.Fields("REFERENC")
                Else
                    FG.Text = ""
                End If
                If cd.Fields("DESCRIP") = "CARGO POR INTERESES" Then
                    FG.Text = Format(cd.Fields("TASA") / 100, "Percent")
                End If
            FG.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FG.Text = cd.Fields("CTABCO")
            Else
                FG.Text = ""
            End If
            If cd.Fields("CVEMOV") > "48" And cd.Fields("CVEMOV") < "60" Then
                  '*Abonos
                If cd.Fields("CVEMOV") = "50" And RENGLON < 2 Then
                    FG.Col = 10
                    FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                    FG.Col = 9
                    FG.Text = ""
                    sdoActual = sdoActual + cd.Fields("IMPORTE")
                Else
                    FG.Col = 9
                    FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                    FG.Col = 10
                    FG.Text = ""
                    sdoActual = sdoActual - cd.Fields("IMPORTE")
                End If
            Else
                '      *Cargos
                FG.Col = 9
                FG.Text = ""
                FG.Col = 10
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual + cd.Fields("IMPORTE")
            End If
            FG.Col = 11
            FG.Text = Format(sdoActual, "Currency")
        End If

    cd.MoveNext
Loop
RenBorra = RENGLON
BorraCeldasenFG
cd.Close

End Sub



Private Sub CmdTrasp_Click()
FlgTrasp = 1
COLOCATITULOSENFG
'MsgBox ("RenBorra=" & RenBorra)
RenBorra = 1
FG.Rows = 2000
BorraCeldasenFG

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM SICMOV ORDER BY NUMES,FECHA,REFERENC,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    numes = 1
    FG.Row = 1
    FG.TopRow = 1
    Do Until cd.EOF = True
        'Or cd.Fields("CTABCO") <> "   "
        If cd.Fields("TIPO") = "T" Then
            If Month(cd.Fields("FECHA")) <> numes Then
                numes = Month(cd.Fields("FECHA"))
                RENGLON = RENGLON + 1
            End If

            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = cd.Fields("NUMREG")
            FG.Col = 1
            FG.Text = cd.Fields("SOCIO")
            FG.Col = 2
            FG.Text = cd.Fields("FECHA")
            FG.Col = 3
            FG.Text = cd.Fields("CVEMOV")
            FG.Col = 4
            FG.Text = cd.Fields("TIPO")
            FG.Col = 5
            FG.Text = cd.Fields("APREPAC")
            FG.Col = 6
            FG.Text = cd.Fields("DESCRIP")
            FG.Col = 7
            If cd.Fields("REFERENC") > 0 Then
                FG.Text = cd.Fields("REFERENC")
            End If
            FG.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FG.Text = cd.Fields("CTABCO")
            End If
            If (cd.Fields("CVEMOV") < "20" Or cd.Fields("CVEMOV") = "48" Or cd.Fields("CVEMOV") = "69") Or (cd.Fields("CVEMOV") > "49" And cd.Fields("CVEMOV") < "59") Then
                    '      *Abonos
                FG.Col = 9
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual + cd.Fields("IMPORTE")
            Else
                '      *Cargos
                FG.Col = 10
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual - cd.Fields("IMPORTE")
            End If
        'Else
        '    If Month(cd.Fields("FECHA")) <> numes Then
        '        numes = Month(cd.Fields("FECHA"))
        '        RENGLON = RENGLON + 1
        '    End If
        '
        '    RENGLON = RENGLON + 1
        '    FG.Col = 0
        '    FG.Row = RENGLON
        '    cd.Fields("NUMREG")
        '    FG.Col = 1
        '    FG.Text = cd.Fields("SOCIO")
        '    FG.Col = 2
        '    FG.Text = cd.Fields("FECHA")
        '    FG.Col = 3
        '    FG.Text = cd.Fields("CVEMOV")
        '    FG.Col = 4
        '    FG.Text = cd.Fields("TIPO")
        '    FG.Col = 5
        '   FG.Text = cd.Fields("APREPAC")

        '    FG.Col = 6
        '    FG.Text = cd.Fields("DESCRIP")
        '    FG.Col = 7
        '    If cd.Fields("REFERENC") > 0 Then
        '        FG.Text = cd.Fields("REFERENC")
        '    End If
        '    FG.Col = 8
        '    If cd.Fields("CTABCO") > 0 Then
        '        FG.Text = cd.Fields("CTABCO")
        '   End If
        '    If (cd.Fields("CVEMOV") < "20" Or cd.Fields("CVEMOV") = "48" Or cd.Fields("CVEMOV") = "69") Or (cd.Fields("CVEMOV") > "49" And cd.Fields("CVEMOV") < "59") Then
        '            '      *Abonos
        '        FG.Col = 9
        '        FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
        '        sdoActual = sdoActual + cd.Fields("IMPORTE")
        '    Else
        '        '      *Cargos
        '        FG.Col = 10
        '        FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
        '        sdoActual = sdoActual - cd.Fields("IMPORTE")
        '    End If
        End If
    cd.MoveNext
Loop
cd.Close
End Sub

Private Sub FG_Click()
        'MsgBox ("FRM AGENDA FEDAMAC-" & frmMiPrimera.Flg)

    If frmMiPrimera.Flg = "1" Then
        FG.Col = 0
        frmMiPrimera.LblSocio = FG.Text

        Static lfrmCount As Long
        Dim frmD As FrmAgenda
        lfrmCount = lfrmCount + 1
        Set frmD = New FrmAgenda
        frmD.Caption = "FrmAgenda"
        frmD.Show
        Exit Sub
    End If
    'MsgBox ("frmMiPrimera.Flg = " & frmMiPrimera.Flg)
    FG.Col = 0
    s_numreg = FG.Text
    FG.Col = 1
    PrvSocio = FG.Text
    FG.Col = 2
    PrvFecha = FG.Text
    TxtModFec = FG.Text
    FG.Col = 3
    PrvCveMov = FG.Text
    FG.Col = 5
    PrvAPrePac = FG.Text
    FG.Col = 6
    PrvDescrip = FG.Text
    TxtModDescrip = FG.Text
    FG.Col = 7
    PrvReferenc = FG.Text
    TxtModRef.Text = FG.Text
    FG.Col = 8

    TxtModBco = FG.Text

    FG.Col = 9
    If FG.Text <> "" Then
        prvImporte = FG.Text
        TxtModImp.Text = FG.Text
    Else
        FG.Col = 10
        If FG.Text <> "" Then
            prvImporte = FG.Text
            TxtModImp.Text = FG.Text
        End If
    End If
    TxtMovSoc = PrvSocio
    'IntRespuesta = MsgBox(prvImporte, 0)
End Sub

Private Sub Form_Load()
    TxtDate = Date
    Flg_movs = 0
    '- 31 - Day(Date)
    'IntRespuesta = MsgBox("Mi Primera Carpeta=" & frmMiPrimera.LblCarpeta, 0)
    Carpeta = frmMiPrimera.LblCarpeta
    LblEmpresa = frmMiPrimera.LblEmpresa
    TxtMovSoc = frmMiPrimera.LblSocio
    'IntRespuesta = MsgBox("Carpeta=" & Carpeta & TxtMovSoc, 0)

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb" '''
    If frmMiPrimera.Flg = "1" Then
        Label3 = "AGENDA FEDAMAC"
        CmdListAgenda_Click
        'frmMiPrimera.Flg = 0

        Exit Sub
    End If
    CmdListMov_Click
End Sub

Private Sub TxtBco_KeyPress(KeyAscii As Integer)
    FlgBco = 1
    If (KeyAscii >= 97) And (KeyAscii <= 122) Then
        KeyAscii = KeyAscii - 32
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtBcoEnter
    End If

End Sub
Private Sub TxtBcoEnter()
COLOCATITULOSENFG


   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM SICMOV ORDER BY CTABCO,FECHA,CVEMOV,REFERENC DESC, IMPORTE", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    numes = 11
    FG.TopRow = 1
    Do Until cd.EOF = True
        UCBco = UCase(TxtBco)
        If cd.Fields("CTABCO") = UCBco Then
            If Month(cd.Fields("FECHA")) <> numes Then
                numes = Month(cd.Fields("FECHA"))
                RENGLON = RENGLON + 1

                FG.Col = 0
                FG.Row = RENGLON

                FG.Text = ""
                FG.Col = 1
                FG.Text = ""
                FG.Col = 2
                FG.Text = ""
                FG.Col = 3
                FG.Text = ""
                FG.Col = 4
                FG.Text = ""
                FG.Col = 5
                FG.Text = ""
                FG.Col = 6
                FG.Text = ""
                FG.Col = 7
                FG.Text = ""
                FG.Col = 8
                FG.Text = ""
                FG.Col = 9
                FG.Text = ""
                FG.Col = 10
                FG.Text = ""
            End If

            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = cd.Fields("NUMREG")
            FG.Col = 1
            FG.Text = cd.Fields("SOCIO")
            FG.Col = 2
            FG.Text = cd.Fields("FECHA")
            FG.Col = 3
            FG.Text = cd.Fields("CVEMOV")
            FG.Col = 4
            FG.Text = cd.Fields("TIPO")
            FG.Col = 5
            FG.Text = cd.Fields("APREPAC")

            FG.Col = 6
            FG.Text = cd.Fields("DESCRIP")
            FG.Col = 7
            If cd.Fields("REFERENC") > 0 Then
                FG.Text = cd.Fields("REFERENC")
            Else
                FG.Text = ""
            End If
            FG.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FG.Text = cd.Fields("CTABCO")
            Else
                FG.Text = ""
            End If
            If cd.Fields("APREPAC") = "A" Or cd.Fields("APREPAC") = "P" Then
                    '      *Abonos
                FG.Col = 9
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual + cd.Fields("IMPORTE")
                FG.Col = 10
                FG.Text = ""
            Else
                '      *Cargos
                FG.Col = 10
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual - cd.Fields("IMPORTE")
                FG.Col = 9
                FG.Text = ""
            End If
        End If

    cd.MoveNext
Loop
cd.Close
RenBorra = RENGLON
BorraCeldasenFG
End Sub



Private Sub TxtMovSoc_Change()
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'ESTA LINEA SOBRA porque ya has especificado el origen en la SELECT:
    'cl.Source = "CLIENTES"        'Especificamos la fuente de datos. En este caso la tabla "CLIENTES"

   'Esto tiene que funcionar
   'cl.Open "SELECT * FROM CLIENTES ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."

   cl.MoveFirst
    'TxtclSocio = pubSocio
    'IntRespuesta = MsgBox(PubSocio, 0)

    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = TxtMovSoc Then
          'TxtclSocio = prvSocio
          'PubSocio = TxtclSocio
          TxtMovNom = cl.Fields("NOMBRE")
          Exit Do
       Else
          TxtMovNom = "No existe nombre de este Socio"

       End If
       cl.MoveNext
    Loop
cl.Close

End Sub
Sub Actualiza_SOCIOS()
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
    cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = TxtMovSoc Then
            'IntRespuesta = MsgBox(PrvAprepac & cl.Fields("SOCIO") & " SOCIO Num:" & PrvSocio & TxtMovSoc & " PrvImporte:" & prvImporte & " TxtModImp" & TxtModImp, 0)
            'IntRespuesta = MsgBox("EL REGISTRO FUE MODIFICADO:-" & PrvAprepac & " SOCIO Num:" & PrvSocio & " Importe:" & TxtModImp, 0)

            SaldoPres = cl.Fields("SALDOPRES")
            Saldo = cl.Fields("SALDO")
            If PrvAPrePac = "P" Then
                cl.Fields("SALDOPRES") = SaldoPres + prvImporte
                cl.Fields("PAGOS") = cl.Fields("PAGOS") - prvImporte
                If Flag_Mod = 1 Then
                    cl.Fields("SALDOPRES") = cl.Fields("SaldoPres") - TxtModImp
                    cl.Fields("PAGOS") = cl.Fields("PAGOS") + TxtModImp
                End If
                cl.Update
                Exit Do
            End If
            If PrvAPrePac = "C" Then
                cl.Fields("SALDOPRES") = SaldoPres - prvImporte
                cl.Fields("PRESTAMOS") = cl.Fields("PRESTAMOS") - prvImporte
                If Flag_Mod = 1 Then
                    cl.Fields("SALDOPRES") = cl.Fields("SaldoPres") + TxtModImp
                    cl.Fields("PRESTAMOS") = cl.Fields("PRESTAMOS") + TxtModImp
                End If
    'IntRespuesta = MsgBox("SOCIOS" & cl.Fields("SALDOPRES") & "-" & cl.Fields("SOCIO"), 1)
                cl.Update
                Exit Do
            End If
            If PrvAPrePac = "A" Then
                cl.Fields("SALDO") = Saldo - prvImporte
                cl.Fields("APORTA") = cl.Fields("APORTA") - prvImporte
                If Flag_Mod = 1 Then
                    cl.Fields("SALDO") = cl.Fields("Saldo") + TxtModImp
                    cl.Fields("APORTA") = cl.Fields("APORTA") + TxtModImp
                End If
                cl.Update
                Exit Do
            End If
            If PrvAPrePac = "R" Then
                cl.Fields("SALDO") = Saldo + prvImporte
                cl.Fields("RETIROS") = cl.Fields("RETIROS") - prvImporte
                If Flag_Mod = 1 Then
                    cl.Fields("SALDO") = cl.Fields("Saldo") - TxtModImp
                    cl.Fields("RETIROS") = cl.Fields("RETIROS") + TxtModImp
                End If
                cl.Update
                Exit Do
            End If
            'cl.Update
            
            'Exit Do
    End If
    cl.MoveNext
    Loop
    Actualiza_99
End Sub
 Sub Actualiza_99()
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
    'DESACTUALIZA CUENTA 99
    cl.MoveFirst
    
    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = "99" Then
            'IntRespuesta = MsgBox("99" & PrvAprepac & cl.Fields("SOCIO") & " SOCIO Num:" & PrvSocio & TxtMovSoc & " PrvImporte:" & prvImporte & " TxtModImp" & TxtModImp, 0)

            If PrvAPrePac = "P" Then
                cl.Fields("SALDO") = cl.Fields("SALDO") - prvImporte
                cl.Fields("APORTA") = cl.Fields("APORTA") - prvImporte
                If Flag_Mod = 1 Then
                    cl.Fields("SALDO") = cl.Fields("Saldo") + TxtModImp
                    cl.Fields("APORTA") = cl.Fields("APORTA") + TxtModImp
                    Flag_Mod = 0
                End If
            End If
            If PrvAPrePac = "C" Then
                cl.Fields("SALDO") = cl.Fields("SALDO") + prvImporte
                cl.Fields("RETIROS") = cl.Fields("RETIROS") - prvImporte
                If Flag_Mod = 1 Then
                    cl.Fields("SALDO") = cl.Fields("Saldo") - TxtModImp
                    cl.Fields("RETIROS") = cl.Fields("RETIROS") + TxtModImp
                    Flag_Mod = 0
                End If
            End If
            If PrvAPrePac = "A" Then
                cl.Fields("SALDO") = cl.Fields("SALDO") - prvImporte
                cl.Fields("APORTA") = cl.Fields("APORTA") - prvImporte
                If Flag_Mod = 1 Then
                    cl.Fields("SALDO") = cl.Fields("Saldo") + TxtModImp
                    cl.Fields("APORTA") = cl.Fields("APORTA") + TxtModImp
                    Flag_Mod = 0
                End If
            End If
            If PrvAPrePac = "R" Then
                cl.Fields("SALDO") = cl.Fields("SALDO") + prvImporte
                cl.Fields("RETIROS") = cl.Fields("RETIROS") - prvImporte
                If Flag_Mod = 1 Then
                    cl.Fields("SALDO") = cl.Fields("Saldo") - TxtModImp
                    cl.Fields("RETIROS") = cl.Fields("RETIROS") + TxtModImp
                    Flag_Mod = 0
                End If
            End If
            cl.Update
            
            Exit Do
    End If
    cl.MoveNext
        
    Loop
    TxtclSocio = ""
End Sub

Private Sub TxtNombre_Change()

'COLOCATITULOSENMS
'Sub COLOCATITULOSENMS()
FG.Row = 0
FG.Col = 0
FG.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
FG.Col = 1
FG.Text = "SOCIO"
FG.Col = 2
FG.Text = "GRUPO"
FG.Col = 3
FG.Text = "NOMBRE"
FG.Col = 4
FG.Text = "SALDO"
FG.Col = 5
FG.Text = "PRESTAMO"
FG.Col = 6
FG.Text = "INTERESES"
FG.Col = 7
FG.Text = "COMISION"

FG.ColWidth(3) = 3000    'AJUSTO EL ANCHO DE LA COLUMNA2
FG.ColWidth(4) = 1000    'AJUSTO EL ANCHO DE LA COLUMNA2
FG.ColWidth(5) = 1000    'AJUSTO EL ANCHO DE LA COLUMNA2
'Label5.Visible = False
'Label4.Visible = False
'TxtCapPres.Visible = False
'TxtCapRet.Visible = False
'TxtCapPres = 0
'TxtCapRet = 0
'totInvGrp = 0
'totPresGrp = 0
'totComGrp = 0
'totIntGrp = 0
'End Sub
'BorraCeldasenFG
'Sub BorraCeldasenFG()
    Do Until RENGLON = 199
       RENGLON = RENGLON + 1
       FG.Col = 0
       FG.Row = RENGLON
       FG.Text = ""
       FG.Col = 1
       FG.Text = ""
       FG.Col = 2
       FG.Text = ""
       FG.Col = 3
       FG.Text = ""
       FG.Col = 4
       FG.Text = ""
       FG.Col = 5
       FG.Text = ""
       FG.Col = 6
       FG.Text = ""
       FG.Col = 7
       FG.Text = ""

    Loop
RENGLON = 0
'FG.Row = 1

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
    LongNom = Len(TxtNombre)
    FG.TopRow = 1
    Do Until cl.EOF = True
        varnombre = Left(cl.Fields("NOMBRE"), LongNom)
        UNOMBRE = TxtNombre
        UNOMBRE = UCase(UNOMBRE)

        If varnombre = UNOMBRE Then
            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = RENGLON
            FG.Col = 1
            FG.Text = cl.Fields("SOCIO")
            FG.Col = 2
            FG.Text = cl.Fields("GRUPO")
            FG.Col = 3
            FG.Text = cl.Fields("NOMBRE")
            FG.Col = 4
            FG.Text = Format(cl.Fields("SALDO"), "Currency")
            FG.Col = 5
            FG.Text = Format(cl.Fields("SALDOPRES"), "Currency")
        End If
        cl.MoveNext
    Loop

cl.Close
'ValorFlexGrid

End Sub
Private Sub CuentaMovs()
    Numovs = 0
    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM SICMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    Do Until cd.EOF = True
        Numovs = Numovs + 1
        UltImporte = cd.Fields("IMPORTE")
        UltSocio = cd.Fields("SOCIO")
        Fecha = cd.Fields("FECHA")
        CveMov = cd.Fields("CVEMOV")
        Tipo = cd.Fields("TIPO")
        Aprepac = cd.Fields("APREPAC")
        CtaBco = cd.Fields("CTABCO")
        Referenc = cd.Fields("REFERENC")
        cd.MoveNext
        Loop
        Numovs = Numovs + 2
        'IntRespuesta = MsgBox("Numovs = " & Numovs, 0)
        cd.Close

End Sub
Private Sub TxtTipomov_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtTipomovEnter
    End If

End Sub

Private Sub TxtTipomovEnter()
COLOCATITULOSENFG
RenBorra = 1
FG.Rows = 2000


   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM SICMOV ORDER BY CTABCO,Id,FECHA,REFERENC DESC,SOCIO,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    FG.TopRow = 1
    Do Until cd.EOF = True
        If cd.Fields("CVEMOV") = TxtTipomov Then
            If Month(cd.Fields("FECHA")) <> numes Then
                numes = Month(cd.Fields("FECHA"))
                RENGLON = RENGLON + 1

                FG.Col = 0
                FG.Row = RENGLON

                FG.Text = ""
                FG.Col = 1
                FG.Text = ""
                FG.Col = 2
                FG.Text = ""
                FG.Col = 3
                FG.Text = ""
                FG.Col = 4
                FG.Text = ""
                FG.Col = 5
                FG.Text = ""
                FG.Col = 6
                FG.Text = ""
                FG.Col = 7
                FG.Text = ""
                FG.Col = 8
                FG.Text = ""
                FG.Col = 9
                FG.Text = ""
                FG.Col = 10
                FG.Text = ""
            End If

            RENGLON = RENGLON + 1
            FG.Col = 0
            FG.Row = RENGLON
            FG.Text = cd.Fields("NUMREG")
            FG.Col = 1
            FG.Text = cd.Fields("SOCIO")
            FG.Col = 2
            FG.Text = cd.Fields("FECHA")
            FG.Col = 3
            FG.Text = cd.Fields("CVEMOV")
            FG.Col = 4
            FG.Text = cd.Fields("TIPO")
            FG.Col = 5
            FG.Text = cd.Fields("APREPAC")

            FG.Col = 6
            FG.Text = cd.Fields("DESCRIP")
            FG.Col = 7
            If cd.Fields("REFERENC") > 0 Then
                FG.Text = cd.Fields("REFERENC")
            Else
                FG.Text = ""
            End If
            FG.Col = 8
            If cd.Fields("CTABCO") > 0 Then
                FG.Text = cd.Fields("CTABCO")
            Else
                FG.Text = ""
            End If
            If cd.Fields("APREPAC") = "A" Or cd.Fields("APREPAC") = "P" Then
                    '      *Abonos
                FG.Col = 9
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual + cd.Fields("IMPORTE")
                FG.Col = 10
                FG.Text = ""
            Else
                '      *Cargos
                FG.Col = 10
                FG.Text = Format(cd.Fields("IMPORTE"), "Currency")
                sdoActual = sdoActual - cd.Fields("IMPORTE")
                FG.Col = 9
                FG.Text = ""
            End If
        End If

    cd.MoveNext
Loop
cd.Close
RenBorra = RENGLON
BorraCeldasenFG
End Sub

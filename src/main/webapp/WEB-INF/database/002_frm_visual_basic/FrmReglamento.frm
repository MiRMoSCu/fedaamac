VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FR 
   Caption         =   "Reglamento"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid FR 
      Height          =   7935
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   13996
      _Version        =   393216
      Rows            =   900
      Cols            =   6
      FixedRows       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "REGLAMENTO INTERNO"
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
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   10335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label REGLAMEN 
      Alignment       =   2  'Center
      Caption         =   " FONDO ECONOMICO DE AYUDA MUTUA, A.C. (FEDAMAC)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ColRow, FlgFR As Integer
Private Carpeta As String


Private Sub Form_Load()

'FR.Width = 10000


    Carpeta = frmMiPrimera.LblCarpeta

FR.ColWidth(0) = 10  'AJUSTO EL ANCHO DE LA COLUMNA2

FR.ColWidth(1) = 10000   'AJUSTO EL ANCHO DE LA COLUMNA2
ColRow = frmMiPrimera.LblSocio
'MsgBox (frmMiPrimera.Flg & " " & ColRow)

If frmMiPrimera.Flg = "1" Then
    ListaReglamento
    frmMiPrimera.Flg = "0"
    Exit Sub
End If
If frmMiPrimera.Flg = "2" Then
    ListaHistoria
    frmMiPrimera.Flg = "0"
    Exit Sub
End If
If frmMiPrimera.Flg = "3" Then
    AgendaDelMes
    frmMiPrimera.Flg = "0"
    Exit Sub
End If
'COLOCATITULOSENMS
'MS.Row = 0
'MS.Col = 0
'MS.Text = ""         'SE TRATA DE NUMERAR LOS REGISTROS
'MS.Col = 1
'FR.Width = 10000

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & "SYSFED" & "\Reglamento.mdb"

   cl.Open "SELECT * FROM REGLAMENTO ORDER BY Id", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
   
    Do Until cl.EOF = True
            RENGLON = RENGLON + 1
            FR.Col = 0
            FR.Row = RENGLON
            FR.Text = cl.Fields("LINEA")
            FR.Col = 1
            FR.Text = cl.Fields("TEXTO")
    cl.MoveNext
Loop
cl.Close


End Sub

Private Sub FR_Click()
 frmMiPrimera.Flg = 1

FR.Col = 0
ColRow = FR.Text
frmMiPrimera.LblSocio = ColRow
'MsgBox (ColRow)
Static lfrmCount As Long
    Dim frmD As FR
    lfrmCount = lfrmCount + 1
    Set frmD = New FR
    frmD.Caption = "FR"
    
    frmD.Show
'ListaReglamento
End Sub
Private Sub ListaReglamento()
'FR.Width = 10000

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & "SYSFED" & "\Reglamento.mdb"

   cl.Open "SELECT * FROM DETALLE ORDER BY Linea,Id", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
    If ColRow = "" Then
        Exit Sub
    End If
    FR.Row = 1
    RENGLON = 0
   FR.TopRow = 0
    Do Until cl.EOF = True
        If cl.Fields("LINEA") > ColRow - 1 Then
            If RENGLON > 778 Then
                Exit Sub
            End If
            RENGLON = RENGLON + 1
            FR.Col = 0
            FR.Row = RENGLON
            FR.Text = RENGLON
            FR.Col = 1
            FR.Text = cl.Fields("TEXTO")
        End If
    cl.MoveNext
Loop
cl.Close


End Sub
Private Sub ListaHistoria()
'FR.Height = 20000

'FR.Width = 20000

FR.ColWidth(0) = 10  'AJUSTO EL ANCHO DE LA COLUMNA2

FR.ColWidth(1) = 900  'AJUSTO EL ANCHO DE LA COLUMNA2
FR.ColWidth(2) = 1000
FR.ColWidth(3) = 2700
FR.ColWidth(4) = 2700
FR.ColWidth(5) = 2700
'FR.ColWidth(6) = 2000
'FR.ColWidth(7) = 2000
'FR.ColWidth(8) = 2000



   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & "SYSFED" & "\Reglamento.mdb"

   cl.Open "SELECT * FROM HISTORIA ORDER BY Id", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
   FR.TopRow = 0
      
    Do Until cl.EOF = True
        'If cl.Fields("LINEA") > ColRow - 1 Then
            RENGLON = RENGLON + 1
            FR.Col = 0
            FR.Row = RENGLON
            FR.Text = RENGLON
            FR.Col = 1
            FR.Text = cl.Fields("CAMPO1")
            FR.Col = 2
            FR.Text = cl.Fields("CAMPO2")
            FR.Col = 3
            FR.Text = cl.Fields("CAMPO3")
            FR.Col = 4
            FR.Text = cl.Fields("CAMPO4")
            FR.Col = 5
            FR.Text = cl.Fields("CAMPO5")
            'FR.Col = 6
            'FR.Text = cl.Fields("CAMPO6")
            'FR.Col = 7
            'FR.Text = cl.Fields("CAMPO7")
            'FR.Col = 8
            'FR.Text = cl.Fields("CAMPO8")

        'End If
    cl.MoveNext
Loop
cl.Close


End Sub
Private Sub AgendaDelMes()
Label2 = "AGENDA FEDAMAC"
'FR.Width = 10000
FR.ColWidth(0) = 10  'AJUSTO EL ANCHO DE LA COLUMNA2

FR.ColWidth(1) = 1200  'AJUSTO EL ANCHO DE LA COLUMNA2
FR.ColWidth(2) = 650
FR.ColWidth(3) = 4000
FR.ColWidth(4) = 2700
FR.ColWidth(5) = 2700
   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & "SYSFED" & "\Agenda.mdb"

   cl.Open "SELECT * FROM AGENDA ORDER BY Fecha", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cl.MoveFirst
   FR.TopRow = 0
    RENGLON = 0
    Cumple = 0
    Do Until cl.EOF = True
        If Month(cl.Fields("FECHA")) = Month(Date) Then
            If Cumple = 0 Then
                FR.Col = 1
                If Date <= cl.Fields("FECHA") Then
                    RENGLON = RENGLON + 1
                    FR.Row = RENGLON
                    FR.Col = 3
                    FR.Text = Format(Date, "ddddd")
                    FR.Col = 4
                    FR.Text = Format(Date, "dddd")
                    Cumple = 1
                End If
            End If
            RENGLON = RENGLON + 1
            FR.Col = 0
            FR.Row = RENGLON
            FR.Text = RENGLON
            FR.Col = 1
            FR.Text = Format(cl.Fields("FECHA"), "ddddd")
            FR.Col = 2
            FR.Text = cl.Fields("HORA") & ":" & cl.Fields("MINUTOS")
            If Date = cl.Fields("FECHA") Then
                FR.Text = "hoy"
            End If
            FR.Col = 3
            FR.Text = cl.Fields("ASUNTO")
            FR.Col = 4
            FR.Text = cl.Fields("UBICA")
            FR.Col = 5
            FR.Text = cl.Fields("INDICA")

        End If
    cl.MoveNext
Loop
cl.Close


End Sub


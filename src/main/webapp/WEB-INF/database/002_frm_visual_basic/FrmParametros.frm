VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmParametros 
   Caption         =   "Parametros"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid FGV 
      Height          =   1935
      Left            =   4080
      TabIndex        =   8
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   9
      Cols            =   3
   End
   Begin VB.TextBox TxtListado 
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Top             =   1800
      Width           =   270
   End
   Begin VB.TextBox TxtFinal 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox TxtFinicial 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox TxtVendedor 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Desglosado Todos (D)"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   " Vendidos (V) Pagados (P)  Cobranza (C)"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Final:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Inicial:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Vendedor:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "FrmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RENGLON As Integer

Private Carpeta As String

Private Sub FGV_Click()
    FGV.Col = 1

    TxtVendedor = FGV.Text
    TxtVendedor.SetFocus
    TxtVendedor.SelStart = Val(TxtVendedor.Text)
End Sub

Private Sub Form_Load()
    Carpeta = frmMiPrimera.LblCarpeta

    TxtVendedor = frmMiPrimera.LblSocio
    TxtFinicial = frmMiPrimera.TxtEjercicio1
    TxtFinal = frmMiPrimera.TxtEjercicio2
    FlgVentas = 1
                  ListaVenDedores

End Sub

Private Sub ListaVenDedores()
FGV.ColWidth(2) = 3000
'TituloVentas
Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SELECT * FROM SOCIOS ORDER BY SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   
    'RENGLON = 0
    cl.MoveFirst
    Do Until cl.EOF = True
        If cl.Fields("TIPO") = "V" Then
            RENGLON = RENGLON + 1
            FGV.Row = RENGLON

            FGV.Col = 0
            FGV.Text = RENGLON

            'Celda = RENGLON
            'BorraCelda
            FGV.Col = 1
            FGV.Text = cl.Fields("SOCIO")
            FGV.Col = 2
            FGV.Text = cl.Fields("NOMBRE")
        End If
    cl.MoveNext
Loop
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub TxtVendedor_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtVendedor.Tag = MODE_OVERTYPE And TxtVendedor.SelLength = 0 Then
        TxtVendedor.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtVendedorEnter

       SendKeys "{tab}"
    End If

End Sub
Private Sub TxtVendedorEnter()
    If TxtVendedor.Tag = MODE_OVERTYPE And TxtVendedor.SelLength = 0 Then
        TxtVendedor.SelLength = 1
    End If

End Sub
Private Sub TxtFinicial_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtFinicial.Tag = MODE_OVERTYPE And TxtFinicial.SelLength = 0 Then
        TxtFinicial.SelLength = 1
    End If

    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtFinicialEnter
       SendKeys "{tab}"
    End If

End Sub
Private Sub TxtFinicialEnter()
    If TxtFinicial.Tag = MODE_OVERTYPE And TxtFinicial.SelLength = 0 Then
        TxtFinicial.SelLength = 1
    End If
    
End Sub

Private Sub TxtFinal_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If TxtFinal.Tag = MODE_OVERTYPE And TxtFinal.SelLength = 0 Then
        TxtFinal.SelLength = 1
    End If

    If KeyAscii = 13 Then
       TxtFinalEnter
        SendKeys "{tab}"
    End If

End Sub
Private Sub TxtFinalEnter()
    If TxtFinal.Tag = MODE_OVERTYPE And TxtFinal.SelLength = 0 Then
        TxtFinal.SelLength = 1
    End If
    
End Sub
Private Sub TxtListado_KeyPress(KeyAscii As Integer)
    ' You have taken some action that changed the text in the
    ' text box. Reset the SelLength if you are in overtype mode.
    If (KeyAscii >= 97) And (KeyAscii <= 122) Then
        KeyAscii = KeyAscii - 32
    End If
    If TxtListado.Tag = MODE_OVERTYPE And TxtListado.SelLength = 0 Then
        TxtListado.SelLength = 1
    End If

    If KeyAscii = 13 Then
        frmMiPrimera.LblSocio = TxtVendedor
        frmMiPrimera.TxtEjercicio2 = TxtFinal
        frmMiPrimera.TxtEjercicio1 = TxtFinicial
        frmMiPrimera.Flg = TxtListado
        
        Static lfrmCount As Long
        Dim frmD As FGP
        lfrmCount = lfrmCount + 1
        Set frmD = New FGP
        frmD.Caption = "FGP"
    
        frmD.Show
        Unload Me
    End If

End Sub

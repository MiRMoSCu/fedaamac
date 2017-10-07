VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FGP 
   Caption         =   "FrmPrestamos"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMorosos 
      Caption         =   "Consulta Socios Morosos"
      Height          =   615
      Left            =   12360
      TabIndex        =   2
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton CmdPrestamos 
      Caption         =   "Consulta Préstamos en el Ejercicio"
      Height          =   735
      Left            =   12360
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid FGP 
      Height          =   7095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12515
      _Version        =   393216
      Rows            =   1000
      Cols            =   9
   End
End
Attribute VB_Name = "FGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Carpeta, Flg As String
Private n_socio As String
Private PrvNombre As String
Private PrvMeses, sihay As Integer
Private PrvPagoMin As String
Private PrvSeguro As String
Private PrvFecVenc As String
Private PrvTasaPres As String
Private PrvSaldoPres As String
Private PrvInteres As String
Private PrvFecPres As String
Private RenBorra As Single
Private meses As Integer
                
Sub Busca_Nombre()

    Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cl As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cl.Open "SOCIOS", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL."

   cl.MoveFirst

    Do Until cl.EOF = True
        If cl.Fields("SOCIO") = n_socio Then
            PrvNombre = cl.Fields("NOMBRE")
            'IntRespuesta = MsgBox(n_socio & ".-" & PrvNombre, 0)
            PrvMeses = (cl.Fields("FECVENC") - cl.Fields("FECPRES")) / 30.4
            meses = PrvMeses
            PrvFecVenc = cl.Fields("FECVENC")
            PrvPagoMin = cl.Fields("PAGOMIN")
            PrvSeguro = cl.Fields("CTASEGURO")
            PrvTasaPres = cl.Fields("TASAPRES")
            PrvSaldoPres = cl.Fields("SALDOPRES")
            PrvInteres = cl.Fields("SALDOPRES") * cl.Fields("TASAPRES") / 100
            PrvFecPres = cl.Fields("FECPRES")
            Exit Do
        End If
        cl.MoveNext
        Loop
        
End Sub

Sub COLOCATITULOSENFGP()
FGP.Row = 0
FGP.Col = 0
FGP.Text = "REG"         'SE TRATA DE NUMERAR LOS REGISTROS
FGP.Col = 1
FGP.Text = "SOCIO"
FGP.Col = 2
FGP.Text = "N O M B R E"
FGP.Col = 3
FGP.Text = "FECHA"
FGP.Col = 4
FGP.Text = "CVE-MOV"
FGP.Col = 5
FGP.Text = "IMPORTE"
FGP.Col = 6
FGP.Text = "DESCRIPCION"
FGP.Col = 7
FGP.Text = "PRIMA"
FGP.Col = 8
FGP.Text = "REFERENCIA"

                    'AJUSTO EL ANCHO DE LAS COLUMNAS
FGP.ColWidth(0) = 500
               
FGP.ColWidth(1) = 600
FGP.ColWidth(2) = 2100
FGP.ColWidth(3) = 1000

FGP.ColWidth(4) = 600
FGP.ColWidth(5) = 1000
FGP.ColWidth(6) = 2000
FGP.ColWidth(7) = 1000

FGP.ColWidth(8) = 1200


End Sub
Sub BorraCeldasenFGP()
RENGLON = RenBorra
    Do Until RENGLON = 999
       RENGLON = RENGLON + 1
       FGP.Col = 0
       FGP.Row = RENGLON
       FGP.Text = ""
       FGP.Col = 1
       FGP.Text = ""
       FGP.Col = 2
       FGP.Text = ""
       FGP.Col = 3
       FGP.Text = ""
       FGP.Col = 4
       FGP.Text = ""
       FGP.Col = 5
       FGP.Text = ""
       FGP.Col = 6
       FGP.Text = ""
       FGP.Col = 7
       FGP.Text = ""
       FGP.Col = 8
       FGP.Text = ""

    Loop

End Sub

Private Sub CmdMorosos_Click()
' Lista de Socios Morosos en una Tabla de FLEX GRID
Dim TotPres As Single
Dim s_grupo As Single
'Dim s_mes As String

COLOCATITULOSENFGP
FGP.ColWidth(1) = 450
FGP.ColWidth(2) = 600
FGP.ColWidth(3) = 3000
FGP.ColWidth(4) = 1000
FGP.ColWidth(5) = 1000
FGP.ColWidth(6) = 1500
FGP.ColWidth(7) = 1300
FGP.ColWidth(8) = 1200

FGP.Col = 1
FGP.Text = "GPO"
FGP.Col = 2
FGP.Text = "SOCIO"
FGP.Col = 3
FGP.Text = "NOMBRE"
FGP.Col = 4
FGP.Text = "Pago Min"
FGP.Col = 5
FGP.Text = "Ultimo Pago"
FGP.Col = 6
FGP.Text = "Saldo Préstamo"
FGP.Col = 7
FGP.Text = "Fecha Préstamo"
FGP.Col = 8
FGP.Text = "Vencimiento"

BorraCeldasenFGP

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   'Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes
   Dim cl As New ADODB.Recordset
    
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   'cd.Open "SELECT * FROM DETMOV ORDER BY GRUPO,SOCIO,FECHA,CVEMOV", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.
   cl.Open "SELECT * FROM SOCIOS ORDER BY GRUPO,SOCIO", cn
    'cd.MoveFirst
    cl.MoveFirst
    Fecorte = cl.Fields("FECORTE")
    s_mes = Month(Fecorte)
    PrimeraLinea = 1
    's_grupo = cd.Fields("GRUPO")

      'SET RELATION TO SOCIO INTO SOCIOS
    Do Until cl.EOF = True

       If cl.Fields("SALDOPRES") > 0.01 Then
        If cl.Fields("PAGOMIN") > cl.Fields("ULTPAGO") Or cl.Fields("ULTPAGO") = 0 Then

          If cl.Fields("GRUPO") < "99" Then
         'If cl.Fields("SOCIO") = "30" Then
         '    IntRespuesta = MsgBox(Month(cl.Fields("FECPRES")) & cl.Fields("socio") & " FGP " & cl.Fields("PAGOMIN") & " " & cl.Fields("ULTPAGO") & s_grupo, 0)
         'End If

           If Month(cl.Fields("FECPRES")) <> Month(Fecorte) Or Year(cl.Fields("FECPRES")) <> Year(Fecorte) Then

            If cl.Fields("GRUPO") <> s_grupo Then
                sihay = 1
                RENGLON = RENGLON + 1
                FGP.Row = RENGLON
                If PrimeraLinea = 1 Then
                    FGP.Col = 3
                    FGP.Text = "SOCIOS MOROSOS"
                    PrimeraLinea = 0
                End If
                s_grupo = cl.Fields("GRUPO")
            End If
            RENGLON = RENGLON + 1
            FGP.Col = 0
            FGP.Row = RENGLON
            FGP.Col = 1
            FGP.Text = cl.Fields("GRUPO")
            FGP.Col = 2
            n_socio = cl.Fields("SOCIO")
            FGP.Text = n_socio
            FGP.Col = 3
            FGP.Text = cl.Fields("NOMBRE")
            FGP.Col = 4
            FGP.Text = Format(cl.Fields("PAGOMIN"), "Currency")
            FGP.Col = 5
            FGP.Text = Format(cl.Fields("ULTPAGO"), "Currency")
            FGP.Col = 6
            FGP.Text = Format(cl.Fields("SALDOPRES"), "Currency")
            FGP.Col = 7
            FGP.Text = cl.Fields("FECPRES")
            FGP.Col = 8
            If cl.Fields("FECVENC") <> "" Then
                FGP.Text = cl.Fields("FECVENC")
            End If
            'RENGLON = RENGLON + 1
            
            n_socio = cl.Fields("SOCIO")
           End If
          End If
        End If
       End If
        cl.MoveNext
        
        Loop
    If sihay < 1 Then
        IntRespuesta = MsgBox("NO EXISTEN SOCIOS MOROSOS", 0)
    End If

End Sub

Private Sub CmdPrestamos_Click()
    'Lista de Préstamos efectuados en el Ejercico en una Tabla de GLEX GRID
Dim TotPres As Single
COLOCATITULOSENFGP

BorraCeldasenFGP

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   Dim cd As New ADODB.Recordset 'Creamos el objeto Recordset.Clientes

   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb"
   
   cd.Open "SELECT * FROM SICMOV ORDER BY NUMES,SOCIO", cn  'Abrimos el Recordset y lo llenamos con la consulta SQL.

   
    cd.MoveFirst
    

    'ABRE_YA ("SOCIOS")
    'ABRE ("SICMOV")
    'SET ORDER TO 1
 
   'SELECT SICMOV
      s_cvemov = cd.Fields("CVEMOV")
      s_tipo = cd.Fields("TIPO")
      m_socio = cd.Fields("SOCIO")

      'SET RELATION TO SOCIO INTO SOCIOS
    Do Until cd.EOF = True
        If cd.Fields("CVEMOV") = "61" And cd.Fields("TIPO") = "B" Or cd.Fields("CVEMOV") = "63" Then
            If n_socio = "" Then
                n_socio = cd.Fields("SOCIO")
            End If
            If cd.Fields("SOCIO") <> n_socio Then
                RENGLON = RENGLON + 1
                FGP.Col = 0
                FGP.Row = RENGLON
                
                FGP.Col = 2
                FGP.Text = "PLAZO: " & meses & " Meses; Vence:"
                FGP.Col = 3
                FGP.Text = PrvFecVenc
                FGP.Col = 4
                FGP.Text = PrvTasaPres & "%"
                FGP.Col = 5
                FGP.Text = Format(TotPres, "Currency")
                FGP.Col = 6
                FGP.Text = "SEGURO DE SOCIOS"
                FGP.Col = 7
                FGP.Text = Format(TotPres * 0.01, "Currency")
                FGP.Col = 8
                FGP.Text = n_socio & "-" & PrvSeguro
                RENGLON = RENGLON + 1
                FGP.Row = RENGLON
                FGP.Col = 0
                FGP.Text = RENGLON
                FGP.Col = 5
                FGP.Text = Format(PrvPagoMin, "Currency")
            
                FGP.Col = 6
                FGP.Text = "PAGO MINIMO"

                RENGLON = RENGLON + 1
                FGP.Col = 0
                FGP.Row = RENGLON
                TotPres = 0
               n_socio = cd.Fields("SOCIO")
            End If

            RENGLON = RENGLON + 1
            FGP.Col = 0
            FGP.Row = RENGLON
            FGP.Text = RENGLON
            FGP.Col = 1
            FGP.Text = cd.Fields("SOCIO")
            
            FGP.Col = 2
            Busca_Nombre
            FGP.Text = PrvNombre
            
            FGP.Col = 3
            FGP.Text = cd.Fields("FECHA")
            FGP.Col = 4
            FGP.Text = cd.Fields("CVEMOV") & "-" & cd.Fields("TIPO")
            FGP.Col = 5
            FGP.Text = Format(cd.Fields("IMPORTE"), "Currency")
            TotPres = TotPres + cd.Fields("IMPORTE")
            FGP.Col = 6
            FGP.Text = cd.Fields("DESCRIP")
            'FGP.Col = 7
            'FGP.Text = Format(cd.Fields("IMPORTE") * 0.01, "Currency")
            FGP.Col = 8
            If cd.Fields("REFERENC") > "" Then
                FGP.Text = cd.Fields("REFERENC")
            End If

        End If
        
    cd.MoveNext
Loop
                RENGLON = RENGLON + 1
                FGP.Col = 0
                FGP.Row = RENGLON

                FGP.Col = 2
                FGP.Text = "PLAZO: " & meses & " Meses; Vence:"
                FGP.Col = 3
                FGP.Text = PrvFecVenc
                FGP.Col = 4
                FGP.Text = PrvTasaPres & "%"
                FGP.Col = 5
                FGP.Text = Format(TotPres, "Currency")
                FGP.Col = 6
                FGP.Text = "SEGURO DE SOCIOS"
                FGP.Col = 7
                FGP.Text = Format(TotPres * 0.01, "Currency")
                FGP.Col = 8
                FGP.Text = n_socio & "-" & PrvSeguro
                RENGLON = RENGLON + 1
                FGP.Row = RENGLON
                FGP.Col = 0
                FGP.Text = RENGLON
                FGP.Col = 5
                FGP.Text = Format(PrvPagoMin, "Currency")
            
                FGP.Col = 6
                FGP.Text = "PAGO MINIMO"
cd.Close
'RenBorra = RENGLON
'BorraCeldasenFGP
End Sub



Private Sub Form_Load()
'    IntRespuesta = MsgBox("Mi Primera Carpeta=" & frmMiPrimera.LblCarpeta, 0)
    Flg = frmMiPrimera.Flg
    Carpeta = frmMiPrimera.LblCarpeta
 '   IntRespuesta = MsgBox("Carpeta=" & Carpeta, 0)

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & Carpeta & "\sisfed.mdb" '''
    CmdMorosos_Click
    If Flg = "1" Then
        CmdPrestamos_Click
        frmMiPrimera.Flg = ""
    End If
End Sub

VERSION 5.00
Begin VB.Form frmMiPrimera 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Aplicación FEDAMAC"
   ClientHeight    =   7020
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   5160
      Picture         =   "frmMiPrimera.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   30
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox Picture7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5160
      Picture         =   "frmMiPrimera.frx":0995
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   28
      Top             =   1080
      Width           =   1095
   End
   Begin VB.PictureBox Picture6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3480
      Picture         =   "frmMiPrimera.frx":167B
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   26
      Top             =   1080
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6600
      Picture         =   "frmMiPrimera.frx":39E6
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   24
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8040
      Picture         =   "frmMiPrimera.frx":4CE0
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   19
      Top             =   2160
      Width           =   1815
      Begin VB.Label Label10 
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   0
         TabIndex        =   21
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1680
      Picture         =   "frmMiPrimera.frx":70CB
      ScaleHeight     =   1995
      ScaleWidth      =   1515
      TabIndex        =   18
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      Picture         =   "frmMiPrimera.frx":8CD8
      ScaleHeight     =   1995
      ScaleWidth      =   1515
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   10080
      Picture         =   "frmMiPrimera.frx":AC6D
      ScaleHeight     =   1635
      ScaleWidth      =   1515
      TabIndex        =   14
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Flg 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   13
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TxtEjercicio2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Text            =   "2012"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox TxtEjercicio1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   10
      Text            =   "2011"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton CmdContPsw 
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtpsw 
      BackColor       =   &H80000004&
      DataField       =   "txtpsw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6240
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "F3D4M4C"
      Top             =   120
      Width           =   1455
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
      Left            =   3120
      TabIndex        =   35
      Top             =   120
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
      Left            =   2760
      TabIndex        =   34
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
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
      Left            =   2160
      TabIndex        =   33
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Contralor 2010-2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   32
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Gerardo Nieto López"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   31
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Jessica Muñoz López Secretaria 2010-2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   29
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Gabriela Carrillo López Tesorera    2010-2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   27
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "Ma de Jesús López Baeza Relaciones Públicas 2010-2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   25
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   23
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "M. Patricia López Baeza Administradora 2010-2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   22
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Julio Nieto Mata Presidente 2010-2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "David López Baeza Socio Fundador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Autor: Luis López Baeza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label LblSocio 
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Ejercicio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label LblCarpeta 
      Alignment       =   2  'Center
      Caption         =   "SYSFED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SYSADMIN V-11.10"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label LblBienvenido 
      Alignment       =   2  'Center
      Caption         =   "BIENVENIDO AL SISTEMA DE ADMINISTRACION DEL FONDO ECONOMICO DE AHORRO Y AYUDA MUTUA, A.C."
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
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   11895
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   6840
      TabIndex        =   5
      Top             =   2760
      Width           =   15
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblpsw 
      AutoSize        =   -1  'True
      Caption         =   "Capture contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1395
   End
   Begin VB.Image ImgCaptura 
      Height          =   345
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "FEDAAMAC"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmMiPrimera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Carpeta As String
Public Provisional As String
Public PubSocio As String
Private flgpsw As Single

Private Sub CmdContPsw_Click()
If txtpsw = "DBGNL" Or txtpsw = "DBVCL" Or txtpsw = "DBJBS" Or txtpsw = "DBGMF" _
        Or txtpsw = "DBBMX" Or txtpsw = "DBLLB" Then
    'ImgCaptura.Picture = LoadPicture("C:\" & Carpeta & "\Happy.bmp")
    'MsgBox ("Sesión: " & txtpsw)
    frmMiPrimera.Flg = "3"
    'Static lfrmCount As Long
    'Dim frmD As FR
    'lfrmCount = lfrmCount + 1
    'Set frmD = New FR
    'frmD.Caption = "FR"
    
    'frmD.Show
    AgendaDelMes
    MsgBox (frmMiPrimera.Flg)
    EmpiezaMENUSYS
    Exit Sub
End If

If txtpsw.Text = "F3D4M4C" Then
    Beep
    lblpsw.Caption = "Contraseña correcta"
    PubSocio = "02"
    frmMiPrimera.Flg = "3"
    'MsgBox (frmMiPrimera.Flg)

    'Static lfrmCount As Long
    'Dim frmD As FR
    'lfrmCount = lfrmCount + 1
    'Set frmD = New FR
    'frmD.Caption = "FR"
    
    'frmD.Show
    AgendaDelMes

    ImgCaptura.Picture = LoadPicture("C:\" & Carpeta & "\Happy.bmp")
    MsgBox ("Contraseña correcta")
        Static lfrmCount As Long
    Dim frmD As frmMENUSYS
    lfrmCount = lfrmCount + 1
    Set frmD = New frmMENUSYS
    frmD.Caption = "frmMENUSYS"
    
    frmD.Show

Else
    lblpsw.Caption = "Contraseña incorrecta"
    txtpsw.Text = ""
    ImgCaptura.Picture = LoadPicture("\" & Carpeta & "\INTL_NO.bmp")
    IntRespuesta = MsgBox("Contraseña INCORRECTA", vbOKCancel)
    
End If
End Sub
Private Sub EmpiezaMENUSYS()
    Static lfrmCount As Long
    Dim frmD As frmMENUSYS
    lfrmCount = lfrmCount + 1
    Set frmD = New frmMENUSYS
    frmD.Caption = "frmMENUSYS"
    
    frmD.Show

End Sub

Private Sub AgendaDelMes()
    Static lfrmCount As Long
    Dim frmD As FR
    lfrmCount = lfrmCount + 1
    Set frmD = New FR
    frmD.Caption = "FR"
    
    frmD.Show

End Sub

Private Sub Form_Load()
    Carpeta = LblCarpeta
    'IntRespuesta = MsgBox("Carpeta=" & Carpeta, 0)
    txtpsw.Tag = MODE_OVERTYPE
    
    Static lfrmCount As Long
    Dim frmD As Mensaje
    lfrmCount = lfrmCount + 1
    Set frmD = New Mensaje
    frmD.Caption = "Mensaje"
    
    frmD.Show

End Sub

Private Sub LblEmpresa_Click()
MsgBox ("Fundado el 31 de Agosto de 1989")
frmMiPrimera.Flg = "2"
Static lfrmCount As Long
    Dim frmD As FR
    lfrmCount = lfrmCount + 1
    Set frmD = New FR
    frmD.Caption = "FR"
    
    frmD.Show

End Sub

Private Sub lblpsw_Click()
    If txtpsw.Text = "F3D4M4C" Then
    Beep
    lblpsw.Caption = "Contraseña correcta"
    ImgCaptura.Picture = LoadPicture("\" & Carpeta & "\Happy.bmp")
    MsgBox ("Contraseña correcta 2")
    Static lfrmCount As Long
    Dim frmD As frmMENUSYS
    lfrmCount = lfrmCount + 1
    Set frmD = New frmMENUSYS
    frmD.Caption = "frmMENUSYS"
    
    frmD.Show
    'FrmSocios.Caption.Show
Else
    lblpsw.Caption = "Contraseña incorrecta"
    txtpsw.Text = ""
    ImgCaptura.Picture = LoadPicture("\" & Carpeta & "\INTL_NO.bmp")
    IntRespuesta = MsgBox("Contraseña INCORRECTA", vbOKCancel)
    
End If
End Sub

Private Sub mnuArchivoSalir_Click()
    End
End Sub

Private Sub mnuConSocios_Click()
    Static lfrmCount As Long
    Dim frmD As FrmSocios
    lfrmCount = lfrmCount + 1
    Set frmD = New FrmSocios
    frmD.Caption = "frmSocios"
    
    frmD.Show
    'FrmSocios.Caption.Show
End Sub

Private Sub mnucaptura_Click()
   If txtpsw.Text = "F3D4M4C" Then
      lblpsw = "Contraseña Correcta"
   'IntRespuesta = MsgBox("Contraseña correcta", 1)

   ImgCaptura.Picture = LoadPicture("\" & Carpeta & "\Happy.bmp")
   lblCapDescrip.Caption = "En esta opción se podrá capturar los movimientos (depósitos, retiros, traspasos, etc.) de cada cuenta de los socios de FEDAMAC"
   lblultmov.Caption = "Ultimo movimiento"
   Txtultimport.Text = "800.00"
   
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
'     TxtrsCtaBco.Text = rsCtaBco

Else

   ImgCaptura.Picture = LoadPicture("\" & Carpeta & "\INTL_NO.bmp")
   lblCapDescrip.Caption = "Contraseña Incorrecta"
   lblultmov.Caption = "Contraseña Incorrecta"
   Txtultimport.Text = ""

End If
End Sub

Private Sub mnuEctaGrupo_Click()
Static lfrmCount As Long
    Dim frmD As MS
    lfrmCount = lfrmCount + 1
    Set frmD = New MS
    frmD.Caption = "MS"
    
    frmD.Show
End Sub

Private Sub mnuListaNombres_Click()
    Static lfrmCount As Long
    Dim frmD As FrmBalance
    lfrmCount = lfrmCount + 1
    Set frmD = New FrmBalance
    frmD.Caption = "frmBalance"
    
    frmD.Show

End Sub

Private Sub mnurelmov_Click()
'IntRespuesta = MsgBox(pbSocio, 0)

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

Private Sub Picture1_Click()
MsgBox ("Nació en Moroleón, Gto. el 25 de Diciembre de 1942")
End Sub

Private Sub Picture2_Click()
MsgBox ("Nació en Moroleón, Gto., el 22 de Septiembre de 1936")
End Sub

Private Sub Picture3_Click()
MsgBox ("Nació en México, D.F. el 10 de Noviembre de 1954")
End Sub

Private Sub Picture4_Click()
MsgBox ("Nació en México, D.F. el 21 de Noviembre de 1962")
End Sub

Private Sub Picture5_Click()
MsgBox ("Nació en México, D.F. el 19 de Noviembre de 1945")
End Sub

Private Sub Picture6_Click()
MsgBox ("Nació en México, D.F., el 19 de Agosto de 1975")
End Sub

Private Sub Picture7_Click()
MsgBox ("Nació en México, D.F., el 25 de Enero de 1983")
End Sub

 Private Sub Txtpsw_KeyPress(KeyAscii As Integer)
    If txtpsw.Tag = MODE_OVERTYPE And txtpsw.SelLength = 0 Then
        txtpsw.SelLength = 1
    End If
    If KeyAscii = 9 Then 'vbKeyTab
        IntRespuesta = MsgBox("KeyAscii=" & KeyAscii & "-" & prvImporte, 0)
    End If
    If KeyAscii = 13 Then
       'IntRespuesta = MsgBox("KeyAscii=13" & KeyAscii & "-" & prvImporte, 0)
       TxtpswEnter


       'SendKeys "{tab}"
    End If
    If flgpsw = 0 Then
        If KeyAscii <> 13 Then
            flgpsw = 1
            txtpsw = ""
        End If
    End If
End Sub
Private Sub TxtpswEnter()
    If txtpsw = "DBGNL" Or txtpsw = "DBVCL" Or txtpsw = "DBJBS" Or txtpsw = "DBGMF" _
        Or txtpsw = "DBBMX" Then
      CmdContPsw_Click
    End If
End Sub


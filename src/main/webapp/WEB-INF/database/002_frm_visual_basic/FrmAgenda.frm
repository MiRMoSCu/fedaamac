VERSION 5.00
Begin VB.Form FrmAgenda 
   Caption         =   "FrmAgenda"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtDateR 
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
      Left            =   1320
      TabIndex        =   30
      Text            =   "25/12/2011"
      Top             =   3000
      Width           =   1095
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   29
      Text            =   "25/12/2011"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox TxtDiaSem 
      Height          =   285
      Left            =   4080
      TabIndex        =   28
      Text            =   "DOMINGO"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox TxtMinutos 
      Height          =   285
      Left            =   3480
      TabIndex        =   27
      Text            =   "00"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox TxtHora 
      Height          =   285
      Left            =   3120
      TabIndex        =   26
      Text            =   "12"
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "SALIR"
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Siguiente Registro"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Registro Anterior"
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Registro en Blanco"
      Height          =   495
      Left            =   3720
      TabIndex        =   20
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Grabar Nuevo Registro"
      Height          =   495
      Left            =   1320
      TabIndex        =   19
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Eiminar Registro"
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton CmdModif 
      Caption         =   "Modificar Datos"
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Siguiente Año"
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Siguiente Mes"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguiente Semana"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente Día"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtIndica 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   " "
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox TxtUbica 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   " "
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox TxtAsunto 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   " "
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label11 
      Caption         =   "Hora:"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Presentar:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Actualizar:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Reprogramar:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "AGENDA FEDAMAC"
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
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label LblRegistro 
      Caption         =   "(Id) (Nuevo)"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Recordatorio:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Indicaciones:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Ubicación:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Asunto:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Num. Registro:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "FrmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PrvDate, PrvDateR As Date
Private Bisiesto, YearEntero, RegAnterior  As Integer
Private cd As New ADODB.Recordset 'Creamos el objeto Recordset.Agenda


Private Sub Command1_Click()
    PrvDate = PrvDate + 1
    PrvDateR = PrvDateR + 1
   TxtDate = PrvDate
   TxtDateR = PrvDateR
End Sub

Private Sub Command10_Click()
            cd.MoveNext
            LblRegistro = cd.Fields("Id")
            PrvDate = cd.Fields("FECHA")
            PrvDateR = cd.Fields("FECHAR")
            TxtDate = cd.Fields("FECHA")
            TxtHora = cd.Fields("HORA")
            TxtMinutos = cd.Fields("MINUTOS")
            TxtAsunto = cd.Fields("ASUNTO")
            TxtUbica = cd.Fields("UBICA")
            TxtIndica = cd.Fields("INDICA")
            TxtDateR = cd.Fields("FECHAR")
End Sub

Private Sub Command2_Click()
    PrvDate = PrvDate + 7
    PrvDateR = PrvDateR + 7
   TxtDate = PrvDate
   TxtDateR = PrvDateR
End Sub

Private Sub Command3_Click()
    If Month(PrvDate) = 4 Or Month(PrvDate) = 6 Or Month(PrvDate) = 9 Or Month(PrvDate) = 11 Then
        PrvDate = PrvDate + 30
    Else
        If Month(PrvDate) = 2 Then
            PrvDate = PrvDate + 28
        Else
            PrvDate = PrvDate + 31
        End If
    End If
    
    If Month(PrvDateR) = 4 Or Month(PrvDateR) = 6 Or Month(PrvDateR) = 9 Or Month(PrvDateR) = 11 Then
        PrvDateR = PrvDateR + 30
    Else
        If Month(PrvDateR) = 2 Then
            PrvDateR = PrvDateR + 28
        Else
            PrvDateR = PrvDateR + 31
        End If
    End If
    TxtDate = PrvDate
    TxtDateR = PrvDateR
End Sub

Private Sub Command4_Click()
    'YearEntero = Year(PrvDate)
    'Bisiesto = YearEntero / 4
    'If Bisiesto * 4 = Year(PrvDate) Then
    '    PrvDate = PrvDate + 366
    'Else
    '    PrvDate = PrvDate + 365
    'End If
   PrvDate = Day(PrvDate) & "/" & Month(PrvDate) & "/" & Year(PrvDate) + 1
 
   PrvDateR = Day(PrvDateR) & "/" & Month(PrvDateR) & "/" & Year(PrvDateR) + 1
   TxtDate = PrvDate
   TxtDateR = PrvDateR
End Sub

Private Sub CmdModif_Click()
Dim retval As Long ' return value


Dim cn As ADODB.Connection
Dim cs As ADODB.Recordset
Dim strPath As String
   
'Update the following path to point to the sample
'Northwind.mdb database on your computer.

strPath = "C:\" & "SYSFED" & "\Agenda.mdb"

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
    .Source = "SELECT * FROM AGENDA"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CursorLocation = adUseServer
    .Open
'End With

  Do Until cs.EOF = True
    'With cs
        If LblRegistro = cs.Fields("Id") Then
            cs.Fields("FECHA") = TxtDate
            cs.Fields("HORA") = TxtHora
            cs.Fields("MINUTOS") = TxtMinutos
            cs.Fields("ASUNTO") = TxtAsunto
            cs.Fields("UBICA") = TxtUbica
            cs.Fields("INDICA") = TxtIndica
            cs.Fields("FECHAR") = TxtDateR
            cs.Update
            Exit Do
        End If
        cs.MoveNext
        
        Loop
End With
    Unload Me
    'MnuAgenda
End Sub

Private Sub Command6_Click()
IntRespuesta = MsgBox("El registro será BORRADO !!!", 1)
 If (IntRespuesta <> 1) Then
    MsgBox ("El registro NO FUE BORRADO !!!")
    Exit Sub
 End If

Dim retval As Long ' return value


Dim cn As ADODB.Connection
Dim cs As ADODB.Recordset
Dim strPath As String
   
'Update the following path to point to the sample
'Northwind.mdb database on your computer.

strPath = "C:\" & "SYSFED" & "\Agenda.mdb"

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
    .Source = "SELECT * FROM AGENDA"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CursorLocation = adUseServer
    .Open
End With
  Do Until cs.EOF = True
    If LblRegistro = cs.Fields("Id") Then
        With cs
        IntRespuesta = MsgBox("El ASUNTO: " & cs.Fields("ASUNTO") & " será BORRADO !!!", 1)
            If (IntRespuesta <> 1) Then
                MsgBox ("El registro NO FUE BORRADO !!!")
            Exit Sub
        End If
        MsgBox (cs.Fields("ASUNTO") & " FUE BORRADO")
        cs.Delete
        cs.Update
        Unload Me
        Exit Sub
        End With
    End If
    cs.MoveNext
        
 Loop
 Unload Me
End Sub

Private Sub Command7_Click()
Dim retval As Long ' return value


Dim cn As ADODB.Connection
Dim cs As ADODB.Recordset
Dim strPath As String
   
'Update the following path to point to the sample
'Northwind.mdb database on your computer.

strPath = "C:\" & "SYSFED" & "\Agenda.mdb"

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
    .Source = "SELECT * FROM AGENDA"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CursorLocation = adUseServer
    .Open
End With

    If LblRegistro = "Nuevo" Then
        With cs

        cs.AddNew
            LblRegistro = cs.Fields("Id")
            cs.Fields("FECHA") = TxtDate
            cs.Fields("HORA") = TxtHora
            cs.Fields("MINUTOS") = TxtMinutos
            cs.Fields("ASUNTO") = TxtAsunto
            cs.Fields("UBICA") = TxtUbica
            cs.Fields("INDICA") = TxtIndica
            cs.Fields("FECHAR") = TxtDateR
            cs.Update
            Unload Me

            Exit Sub
        End With
    Else
        MsgBox ("Este registro YA EXISTE")
    End If

    Unload Me
End Sub

Private Sub Command8_Click()
    LblRegistro = "Nuevo"
    TxtDate = Date
    TxtHora = "'"
    TxtMinutos = "'"
    TxtAsunto = "'"
    TxtUbica = "'"
    TxtIndica = "'"
    TxtDateR = Date
End Sub

Private Sub Command9_Click()
    cd.MoveFirst
    Do Until cd.EOF = True
        If cd.Fields("Id") = RegAnterior Then
            LblRegistro = cd.Fields("Id")
            PrvDate = cd.Fields("FECHA")
            PrvDateR = cd.Fields("FECHAR")
            TxtDate = cd.Fields("FECHA")
            TxtHora = cd.Fields("HORA")
            TxtMinutos = cd.Fields("MINUTOS")
            TxtAsunto = cd.Fields("ASUNTO")
            TxtUbica = cd.Fields("UBICA")
            TxtIndica = cd.Fields("INDICA")
            TxtDateR = cd.Fields("FECHAR")
            Exit Sub
        End If
        'RegAnterior = cd.Fields("Id")
        cd.MoveNext
        Loop
End Sub

Private Sub Form_Load()
    LblRegistro = frmMiPrimera.LblSocio

    LblEmpresa = frmMiPrimera.LblEmpresa
    TxtMovSoc = frmMiPrimera.LblSocio
    'IntRespuesta = MsgBox("Carpeta=" & Carpeta & TxtMovSoc, 0)

   Dim cn As New ADODB.Connection        'Creamos el objeto Connection

   
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=c:\" & "SYSFED" & "\Agenda.mdb" '''
   
   cd.Open "SELECT * FROM AGENDA ORDER BY FECHA", cn  'Abrimos el Recordset y lo llenamos con la consulta S
   
   cd.MoveFirst
    
    Do Until cd.EOF = True
        If cd.Fields("Id") = LblRegistro Then
            PrvDate = cd.Fields("FECHA")
            PrvDateR = cd.Fields("FECHAR")
            TxtDate = cd.Fields("FECHA")
            TxtMinutos = cd.Fields("MINUTOS")
            TxtHora = cd.Fields("HORA")
            TxtDiaSem = Format(cd.Fields("FECHA"), "dddd")
            TxtAsunto = cd.Fields("ASUNTO")
            TxtUbica = cd.Fields("UBICA")
            TxtIndica = cd.Fields("INDICA")
            TxtDateR = cd.Fields("FECHAR")
            Exit Do
        End If
        RegAnterior = cd.Fields("Id")
        cd.MoveNext
        Loop
End Sub
Private Sub MnuAgenda()
Unload Me

frmMiPrimera.Flg = "1"

Static lfrmCount As Long
    Dim frmD As FG
    lfrmCount = lfrmCount + 1
    Set frmD = New FG
    frmD.Caption = "FG"
    
    frmD.Show
End Sub

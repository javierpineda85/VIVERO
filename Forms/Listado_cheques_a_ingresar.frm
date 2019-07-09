VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Listado_cheques_a_ingresar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de cheques a ingresar al banco"
   ClientHeight    =   9360
   ClientLeft      =   780
   ClientTop       =   855
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Listado_cheques_a_ingresar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13470
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      Picture         =   "Listado_cheques_a_ingresar.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   2880
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Listar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6480
         Picture         =   "Listado_cheques_a_ingresar.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Listado_cheques_a_ingresar.frx":109E
         DragIcon        =   "Listado_cheques_a_ingresar.frx":1628
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5400
         Picture         =   "Listado_cheques_a_ingresar.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61472769
         CurrentDate     =   41087
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61472769
         CurrentDate     =   40909
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Buscar desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   14
      Top             =   1440
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fecha"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      DisabledPicture =   "Listado_cheques_a_ingresar.frx":213C
      DragIcon        =   "Listado_cheques_a_ingresar.frx":26C6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      Picture         =   "Listado_cheques_a_ingresar.frx":2C50
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Listar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      Picture         =   "Listado_cheques_a_ingresar.frx":31DA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5640
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6135
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   10821
      _Version        =   393216
      BackColor       =   16777152
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Cheques por Ingresar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   13
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar por Nº de cheque:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Monto Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   10
      Top             =   8880
      Width           =   3855
   End
End
Attribute VB_Name = "Listado_cheques_a_ingresar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
a = "select * from salecheque order by vto"
TABLA.Open a, conexion_BD
MSFlexGrid1.Clear
Call Alta_cheque

With MSFlexGrid1
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!n_interno
    .TextMatrix(lin, 1) = TABLA!n_cheque
    .TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    .TextMatrix(lin, 3) = TABLA!banco
    .TextMatrix(lin, 4) = TABLA!vto
    .TextMatrix(lin, 5) = TABLA!detalle
    .TextMatrix(lin, 6) = TABLA!destino

    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
    
    TABLA.MoveNext
Loop
End With
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub Command2_Click()
Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)

LISTADO = "select * from salecheque where vto >= " + CStr(inicio) + " and vto <= " + CStr(final) + " order by vto"
TABLA.Open LISTADO, conexion_BD

MSFlexGrid1.Clear

With MSFlexGrid1
Call Alta_cheque
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!n_interno
    .TextMatrix(lin, 1) = TABLA!n_cheque
    .TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    .TextMatrix(lin, 3) = TABLA!banco
    .TextMatrix(lin, 4) = TABLA!vto
    .TextMatrix(lin, 5) = TABLA!detalle
    .TextMatrix(lin, 6) = TABLA!destino

    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        'MsgBox " Este cheque ya ha sido ingresado al banco"
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
    
    TABLA.MoveNext
Loop
End With
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub Command3_Click()
a = "select * from salecheque order by vto"
TABLA.Open a, conexion_BD
MSFlexGrid1.Clear

With MSFlexGrid1
Call Alta_cheque
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = MSFlexGrid1.Rows + 1
    .TextMatrix(lin, 0) = TABLA!n_interno
    .TextMatrix(lin, 1) = TABLA!n_cheque
    .TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    .TextMatrix(lin, 3) = TABLA!banco
    .TextMatrix(lin, 4) = TABLA!vto
    .TextMatrix(lin, 5) = TABLA!detalle
    .TextMatrix(lin, 6) = TABLA!destino

    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
    
    TABLA.MoveNext
Loop
End With
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub Command4_Click()

    a = "select * from salecheque where n_cheque = " & Val(Text1) & " order by vto"
    TABLA.Open a, conexion_BD
MSFlexGrid1.Clear
    Call Alta_cheque
    
    With MSFlexGrid1
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!n_interno
    .TextMatrix(lin, 1) = TABLA!n_cheque
    .TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    .TextMatrix(lin, 3) = TABLA!banco
    .TextMatrix(lin, 4) = TABLA!vto
    .TextMatrix(lin, 5) = TABLA!detalle
    .TextMatrix(lin, 6) = TABLA!destino

    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        MsgBox " Este cheque ya ha sido ingresado al banco"
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
    
    TABLA.MoveNext
Loop
End With
Label3 = Format(importe, "currency")
TABLA.Close

End Sub

Private Sub Command5_Click()
Printer.Orientation = 1
Printer.FontSize = 10
Printer.Font = arial
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Fecha: "; Date; Tab(55); SISTEMA; Tab(90); Page
Printer.Print Tab(50); SISTEMA_DIR
'Printer.Print Tab(10); "Fecha: "; Date; Tab(55); " STELLA DAVIRE"; Tab(90); Page
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(30); "LISTADO DE CHEQUES POR INGRESAR"
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Detalle desde "; DTPicker1; " Hasta "; DTPicker2
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.FontSize = 8
Printer.Print Tab(10); "Nº INTERNO"; Tab(25); "Nº CHEQUE"; Tab(40); "IMPORTE"; Tab(57); "BANCO"; Tab(67); "FECHA VTO"; Tab(95); "DETALLE"; Tab(130); "DESTINO"
Printer.Print Tab(10); "==============================================================================================================="

For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    interno = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    cheque = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    importe = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    banco = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
    vto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4)
    detalle = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 5)
    destino = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 6)

Printer.Print Tab(12); interno; Tab(27); cheque; Tab(42); importe; Tab(55); banco; Tab(69); vto; Tab(85); detalle; Tab(125); destino
Printer.Print Tab(10); "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"


If i = 32 Or i = 68 Or i = 103 Or i = 138 Then
    Printer.NewPage
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); "Nº INTERNO"; Tab(25); "Nº CHEQUE"; Tab(40); "IMPORTE"; Tab(57); "BANCO"; Tab(67); "FECHA VTO"; Tab(95); "DETALLE"; Tab(130); "DESTINO"
    Printer.Print Tab(10); "==============================================================================================================="
End If

Next
Printer.Print Tab(10); ""
Printer.Print Tab(10); " TOTAL:"; Label3
Printer.EndDoc

End Sub

Private Sub Form_Load()


DTPicker2 = Date

a = "select * from salecheque order by vto asc"
TABLA.Open a, conexion_BD
With MSFlexGrid1
Call Alta_cheque
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!n_interno
    .TextMatrix(lin, 1) = TABLA!n_cheque
    .TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    .TextMatrix(lin, 3) = TABLA!banco
    .TextMatrix(lin, 4) = TABLA!vto
    .TextMatrix(lin, 5) = TABLA!detalle
    .TextMatrix(lin, 6) = TABLA!destino

    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
    

    TABLA.MoveNext
Loop
End With
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub MSFlexGrid1_Click()
datos = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0) ' datos es variable general y .row,0
                                                         ' es el numero de interno
' selecciona los datos del cheque por nº de interno y lo manda a datos_a_banco para realizar el proceso
Datos_a_banco.Show
End Sub

Private Sub Alta_cheque()

With MSFlexGrid1
.FixedCols = 0
.Cols = 7
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = " Nº INTERNO"
.TextMatrix(0, 1) = "Nº CHEQUE"
.TextMatrix(0, 2) = "IMPORTE"
.TextMatrix(0, 3) = "BANCO"
.TextMatrix(0, 4) = "VENCIMIENTO"
.TextMatrix(0, 5) = "DETALLE"
.TextMatrix(0, 6) = "DESTINO"

.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 3000
.ColWidth(6) = 3000

End With
End Sub

Private Sub Option1_Click()
If Option1 = True Then
    Frame2.BorderStyle = 0
    Frame2.Visible = True
    Label1.Visible = False
    Text1.Visible = False
    Command1.Visible = False
    Command4.Visible = False
End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
    Label1.Visible = True
    Text1.Visible = True
    Command1.Visible = True
    Command4.Visible = True
    Frame2.Visible = False
End If

End Sub

Private Sub Text1_Click()
If KeyAscii = 13 Then
    a = "select * from salecheque where n_cheque = " & Val(Text1) & " order by vto"
    TABLA.Open a, conexion_BD
    MSFlexGrid1.Clear
    Call Alta_cheque
    
    With MSFlexGrid1
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!n_interno
    .TextMatrix(lin, 1) = TABLA!n_cheque
    .TextMatrix(lin, 2) = TABLA!vto
    .TextMatrix(lin, 3) = TABLA!detalle
    .TextMatrix(lin, 4) = Format(TABLA!importe, "currency")
    .TextMatrix(lin, 5) = TABLA!banco
    .TextMatrix(lin, 6) = TABLA!destino

    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        MsgBox " Este cheque ya ha sido ingresado al banco"
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
    
    TABLA.MoveNext
Loop
    End With
Label3 = importe
TABLA.Close
End If
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Listados_cheques_en_banco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados de Cheques Emitidos"
   ClientHeight    =   9450
   ClientLeft      =   885
   ClientTop       =   915
   ClientWidth     =   11400
   Icon            =   "Listados_cheques_en_banco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   11400
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
      Left            =   10200
      Picture         =   "Listados_cheques_en_banco.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   2640
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Listados_cheques_en_banco.frx":0B14
         DragIcon        =   "Listados_cheques_en_banco.frx":109E
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
         Picture         =   "Listados_cheques_en_banco.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
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
         Left            =   6480
         Picture         =   "Listados_cheques_en_banco.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
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
         Height          =   360
         Left            =   2640
         TabIndex        =   6
         Top             =   420
         Width           =   1815
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
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   2415
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
      Left            =   240
      TabIndex        =   16
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
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   2640
      TabIndex        =   12
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
         Picture         =   "Listados_cheques_en_banco.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Listados_cheques_en_banco.frx":26C6
         DragIcon        =   "Listados_cheques_en_banco.frx":2C50
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
         Picture         =   "Listados_cheques_en_banco.frx":31DA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3840
         TabIndex        =   4
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
         Format          =   61603841
         CurrentDate     =   41087
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
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
         Format          =   61603841
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6015
      Left            =   600
      TabIndex        =   9
      Top             =   2520
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10610
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
      Left            =   5640
      TabIndex        =   11
      Top             =   8880
      Width           =   4215
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
      Left            =   3360
      TabIndex        =   10
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Cheques en Banco"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   6375
   End
End
Attribute VB_Name = "Listados_cheques_en_banco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
a = "select * from enbanco order by vto desc"
TABLA.Open a, conexion_BD
MSFlexGrid1.Clear
lin = 0
Call Alta_cheque
Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = TABLA!vto
    MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_ingreso
    
    importe = CDbl(importe) + TABLA!importe
   
    TABLA.MoveNext
Loop
    
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub Command2_Click()
Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)

LISTADO = "select * from enbanco where vto >= " + CStr(inicio) + " and vto <= " + CStr(final) + " order by vto"
TABLA.Open LISTADO, conexion_BD

MSFlexGrid1.Clear
lin = 0
Call Alta_cheque
Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = TABLA!vto
    MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_ingreso
    
    importe = CDbl(importe) + TABLA!importe
   
    TABLA.MoveNext
Loop
    
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub Command3_Click()
a = "select * from enbanco order by vto desc"
TABLA.Open a, conexion_BD
MSFlexGrid1.Clear
lin = 0
Call Alta_cheque
Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = TABLA!vto
    MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_ingreso
    
    importe = CDbl(importe) + TABLA!importe
   
    TABLA.MoveNext
Loop
    
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub Command4_Click()
a = "select * from enbanco where n_cheque = " & Val(Text1) & " order by vto"
TABLA.Open a, conexion_BD
MSFlexGrid1.Clear
lin = 0
Call Alta_cheque
Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = TABLA!vto
    MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_ingreso
    
    importe = CDbl(importe) + TABLA!importe
   
    TABLA.MoveNext
Loop
    
Label3 = Format(importe, "currency")
TABLA.Close

End Sub

Private Sub Command5_Click()
Printer.Orientation = 1
Printer.FontSize = 10
Printer.Font = arial
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Fecha: "; Date; Tab(59); SISTEMA; Tab(90); Page
Printer.Print Tab(50); SISTEMA_DIR
'Printer.Print Tab(10); "Fecha: "; Date; Tab(59); " STELLA DAVIRE"; Tab(90); Page
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(30); "LISTADO DE CHEQUES EN EL BANCO"
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Detalle desde "; DTPicker1; " Hasta "; DTPicker2
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.FontSize = 8
Printer.Print Tab(10); "Nº INTERNO"; Tab(25); "Nº CHEQUE"; Tab(40); "FECHA VTO"; Tab(57); "IMPORTE"; Tab(73); "FECHA INGRESO" ' Tab(95); "DETALLE"; Tab(130); "DESTINO"
Printer.Print Tab(10); "================================================================"

For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    inter = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    che = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    venc = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    impor = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
    ingr = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4)

Printer.Print Tab(12); inter; Tab(27); che; Tab(41); venc; Tab(59); impor; Tab(75); ingr; 'Tab(85); detalle; Tab(125); destino
Printer.Print Tab(10); "-----------------------------------------------------------------------------------------------------------------"
If i = 32 Or i = 68 Or i = 103 Or i = 138 Then
    Printer.NewPage
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); "Nº INTERNO"; Tab(25); "Nº CHEQUE"; Tab(40); "FECHA VTO"; Tab(57); "IMPORTE"; Tab(73); "FECHA INGRESO" ' Tab(95); "DETALLE"; Tab(130); "DESTINO"
    Printer.Print Tab(10); "================================================================"
End If
Next

Printer.Print Tab(10); ""
Printer.Print Tab(10); " TOTAL: "; Label3

Printer.EndDoc

End Sub

Private Sub Form_Load()
DTPicker2 = Date

a = "select * from enbanco order by vto desc"
TABLA.Open a, conexion_BD

Call Alta_cheque
Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = TABLA!vto
    MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_ingreso
    
    If TABLA!importe < 0 Then
        importe = CDbl(importe) + TABLA!importe * -1  ' SUMA LOS IMPORTES
        MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!importe * -1, "currency") ' DEJA EN POSITIVO LOS VALORES
    Else
        importe = CDbl(importe) + TABLA!importe
    End If
    
    TABLA.MoveNext
Loop
    
Label3 = Format(importe, "currency")
TABLA.Close
End Sub


Private Sub Alta_cheque()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 5
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "Nº INTERNO"
MSFlexGrid1.TextMatrix(0, 1) = "Nº CHEQUE"
MSFlexGrid1.TextMatrix(0, 2) = "VENCIMIENTO"
MSFlexGrid1.TextMatrix(0, 3) = "IMPORTE"
MSFlexGrid1.TextMatrix(0, 4) = "FECHA INGRESO"


MSFlexGrid1.ColWidth(0) = 2000
MSFlexGrid1.ColWidth(1) = 2000
MSFlexGrid1.ColWidth(2) = 2000
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 2000

End Sub

Private Sub Option1_Click()
If Option1 = True Then
    Frame2.Visible = True
    Frame2.BorderStyle = 0
    Frame3.Visible = False
Else
    
    Frame3.Visible = True
    Frame3.BorderStyle = 0
    Frame2.Visible = False
    
End If

End Sub

Private Sub Option2_Click()
If Option2 = True Then
    Frame3.Visible = True
    Frame3.BorderStyle = 0
    Frame2.Visible = False
Else
    
    Frame2.Visible = True
    Frame2.BorderStyle = 0
    Frame3.Visible = False
    
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

a = "select * from enbanco where n_cheque = " & Val(Text1) & " order by vto"
TABLA.Open a, conexion_BD
MSFlexGrid1.Clear
lin = 0
Call Alta_cheque
Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = TABLA!vto
    MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_ingreso
    
    importe = CDbl(importe) + TABLA!importe
   
    TABLA.MoveNext
Loop
    
Label3 = Format(importe, "currency")
TABLA.Close

End If
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Listado_retenciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de reteciones"
   ClientHeight    =   7755
   ClientLeft      =   1380
   ClientTop       =   915
   ClientWidth     =   10350
   Icon            =   "Listado_retenciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10350
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
      Left            =   7200
      Picture         =   "Listado_retenciones.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
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
      Left            =   6120
      Picture         =   "Listado_retenciones.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22216705
      CurrentDate     =   41058
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22216705
      CurrentDate     =   40909
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4935
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8705
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6720
      TabIndex        =   8
      Top             =   7080
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE DE RETENCIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "SALDO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   7200
      Width           =   975
   End
End
Attribute VB_Name = "Listado_retenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)

LISTADO = "select * from retencion where fecha >= " + CStr(inicio) + " and fecha <= " + CStr(final) + " order by fecha desc"
TABLA.Open LISTADO, conexion_BD

MSFlexGrid1.Clear
Call ALTAGRID
saldo = 0
saldo1 = 0
resta = 0

    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!destino
        MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
        MSFlexGrid1.TextMatrix(lin, 3) = TABLA!r_interno
        saldo = CDbl(saldo) + TABLA!importe
        
        If TABLA!importe <= 0 Then
            MSFlexGrid1.Row = lin
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
            lin = lin - 1
        End If
        
        TABLA.MoveNext
    Loop
TABLA.Close

Label5.BackColor = vbGreen
Label5.ForeColor = vbBlack
Label5 = Format(saldo, "currency")


End Sub

Private Sub Command5_Click()
Printer.Orientation = 1
Printer.FontSize = 10
Printer.Font = arial
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Fecha: "; Date; Tab(59); SISTEMA;
Printer.Print Tab(50); SISTEMA_DIR
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(30); "LISTADO DE RETENCIONES"
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Detalle desde "; DTPicker1; " Hasta "; DTPicker2
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.FontSize = 8
Printer.Print Tab(12); "FECHA"; Tab(35); "DESTINO"; Tab(75); "IMPORTE"; Tab(90); "REMITO" '; Tab(83); "CLIENTE"; Tab(120); "FECHA VTO"; Tab(137); "FECHA MODIF"
Printer.Print Tab(10); "======================================================================"

For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    fecha = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    destino = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    importe = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    r_interno = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)


Printer.Print Tab(10); fecha; Tab(25); destino; Tab(77); importe; Tab(92); r_interno; ' Tab(76); cliente; Tab(122); vto; Tab(139); depo
Printer.Print Tab(10); "----------------------------------------------------------------------------------------------------------------------------"

If i = 32 Or i = 68 Or i = 103 Or i = 138 Then
    Printer.NewPage
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(12); "FECHA"; Tab(35); "DESTINO"; Tab(75); "IMPORTE"; Tab(90); "REMITO" '; Tab(83); "CLIENTE"; Tab(120); "FECHA VTO"; Tab(137); "FECHA MODIF"
    Printer.Print Tab(10); "======================================================================"
    
End If

Next
Printer.Print Tab(10); ""
Printer.Print Tab(10); " TOTAL:"; Label5
Printer.EndDoc
End Sub

Private Sub Form_Load()

DTPicker2 = Date
Call ALTAGRID
LISTADO = "select * from retencion order by fecha desc"
TABLA.Open LISTADO, conexion_BD
saldo = 0
saldo1 = 0
resta = 0

    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!destino
        MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
        MSFlexGrid1.TextMatrix(lin, 3) = TABLA!r_interno
        saldo = CDbl(saldo) + TABLA!importe
        
        If TABLA!importe <= 0 Then
            MSFlexGrid1.Row = lin
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
            lin = lin - 1
        End If
        TABLA.MoveNext
    Loop
TABLA.Close

Label5.BackColor = vbGreen
Label5.ForeColor = vbBlack
Label5 = Format(saldo, "currency")
End Sub
Private Sub ALTAGRID()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 4
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "FECHA"
MSFlexGrid1.TextMatrix(0, 1) = "NOMBRE/RAZON SOCIAL"
MSFlexGrid1.TextMatrix(0, 2) = "IMPORTE"
MSFlexGrid1.TextMatrix(0, 3) = "Nº DE REG"

MSFlexGrid1.ColWidth(0) = 1800
MSFlexGrid1.ColWidth(1) = 4900
MSFlexGrid1.ColWidth(2) = 1400
MSFlexGrid1.ColWidth(3) = 1400

End Sub


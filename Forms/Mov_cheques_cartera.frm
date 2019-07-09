VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mov_cheques_cartera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheques en cartera"
   ClientHeight    =   9360
   ClientLeft      =   885
   ClientTop       =   960
   ClientWidth     =   13470
   Icon            =   "Mov_cheques_cartera.frx":0000
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
      Picture         =   "Mov_cheques_cartera.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   3000
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Mov_cheques_cartera.frx":0B14
         DragIcon        =   "Mov_cheques_cartera.frx":109E
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
         Picture         =   "Mov_cheques_cartera.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
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
         Picture         =   "Mov_cheques_cartera.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
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
         Format          =   22282241
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
         Format          =   22282241
         CurrentDate     =   40909
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
         TabIndex        =   18
         Top             =   480
         Width           =   1335
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
      Left            =   600
      TabIndex        =   15
      Top             =   1560
      Width           =   2415
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
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   3000
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   7575
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
         TabIndex        =   3
         Top             =   420
         Width           =   1815
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
         Picture         =   "Mov_cheques_cartera.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Mov_cheques_cartera.frx":26C6
         DragIcon        =   "Mov_cheques_cartera.frx":2C50
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
         Picture         =   "Mov_cheques_cartera.frx":31DA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label4 
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
         TabIndex        =   14
         Top             =   480
         Width           =   2415
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5655
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9975
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
      BackStyle       =   0  'Transparent
      Caption         =   "saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8640
      TabIndex        =   12
      Top             =   8640
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DE CHEQUES EN CARTERA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Mov_cheques_cartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
P = "select * from entracheque order by fecha_vto"
TABLA.Open P, conexion_BD
Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
   
    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        MSFlexGrid1.Row = lin
        MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
        lin = lin - 1
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If

    
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

LISTADO = "select * from entracheque where fecha_vto >= " + CStr(inicio) + " and fecha_vto <= " + CStr(final) + " order by fecha_vto"
TABLA.Open LISTADO, conexion_BD

MSFlexGrid1.Clear

Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
   
    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        MSFlexGrid1.Row = lin
        MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
        lin = lin - 1
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
     
    
    TABLA.MoveNext
Loop
Label3 = Format(importe, "currency")
TABLA.Close


End Sub

Private Sub Command3_Click()
P = "select * from entracheque order by fecha_vto"
TABLA.Open P, conexion_BD
Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
   
    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        MSFlexGrid1.Row = lin
        MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
        lin = lin - 1
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
     
    
    TABLA.MoveNext
Loop
Label3 = Format(importe, "currency")
TABLA.Close

End Sub

Private Sub Command4_Click()
a = "select * from entracheque where n_cheque = " & Val(Text1) & " order by fecha_vto"
TABLA.Open a, conexion_BD
MSFlexGrid1.Clear
Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
   
    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        MSFlexGrid1.Row = lin
        MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
        lin = lin - 1
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If

    
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
Printer.Print Tab(10); "Fecha: "; Date; Tab(55); SISTEMA; Tab(90); Page
Printer.Print Tab(50); SISTEMA_DIR
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(30); "LISTADO DE CHEQUES EN CARTERA"
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Detalle desde "; DTPicker1; " Hasta "; DTPicker2
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.FontSize = 8

Printer.Print Tab(10); "Nº INTERNO"; Tab(25); "Nº CHEQUE"; Tab(43); "IMPORTE"; Tab(60); "BANCO"; Tab(83); "CLIENTE"; Tab(120); "FECHA VTO"; Tab(137); "FECHA MODIF"
Printer.Print Tab(10); "================================================================================================================================================================"

For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    interno = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    cheque = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    importe = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    banco = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
    cliente = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4)
    vto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 5)
    depo = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 6)


Printer.Print Tab(12); interno; Tab(27); cheque; Tab(45); importe; Tab(57); banco; Tab(76); cliente; Tab(122); vto; Tab(139); depo
Printer.Print Tab(10); "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Next
Printer.Print Tab(10); ""
Printer.Print Tab(10); " TOTAL: "; Label3
Printer.EndDoc
End Sub

Private Sub Form_Load()
DTPicker2 = Date

P = "select * from entracheque order by fecha_vto"
TABLA.Open P, conexion_BD
Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
    
    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        MSFlexGrid1.Row = lin
        MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
        lin = lin - 1
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
    

    
    TABLA.MoveNext
Loop
Label3 = Format(importe, "currency")
TABLA.Close

End Sub
Private Sub Alta_Cheques()
With MSFlexGrid1
.FixedCols = 0
.Cols = 7
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "Nº INTERNO"
.TextMatrix(0, 1) = "Nº DE CHEQUE"
.TextMatrix(0, 2) = "IMPORTE"
.TextMatrix(0, 3) = "BANCO"
.TextMatrix(0, 4) = "CLIENTE"
.TextMatrix(0, 5) = "FECHA DE VTO"
.TextMatrix(0, 6) = "FECHA MODIF."


.ColWidth(0) = 1000
.ColWidth(1) = 1700
.ColWidth(2) = 1500
.ColWidth(3) = 2000
.ColWidth(4) = 3000
.ColWidth(5) = 1500
.ColWidth(6) = 1500

End With
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

a = "select * from entracheque where n_cheque = " & Val(Text1) & " order by fecha_vto"
TABLA.Open a, conexion_BD
MSFlexGrid1.Clear
Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
   
    If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
        MSFlexGrid1.Row = lin
        MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
        lin = lin - 1
    Else
    
        importe = CDbl(importe) + TABLA!importe
        
    End If
     
    TABLA.MoveNext
Loop
Label3 = Format(importe, "currency")
TABLA.Close
End If
End Sub

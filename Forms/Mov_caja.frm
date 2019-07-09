VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mov_caja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de caja"
   ClientHeight    =   8160
   ClientLeft      =   2295
   ClientTop       =   960
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mov_caja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7800
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "x fecha"
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
      Left            =   3000
      Picture         =   "Mov_caja.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4320
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   420
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      ItemData        =   "Mov_caja.frx":0B14
      Left            =   1440
      List            =   "Mov_caja.frx":0B24
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Height          =   345
      Left            =   4920
      TabIndex        =   18
      Top             =   7560
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
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
      Left            =   1680
      Picture         =   "Mov_caja.frx":0B51
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2295
      Left            =   360
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   16777152
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22216705
      CurrentDate     =   41023
      MinDate         =   2
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
      _ExtentX        =   3413
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
      CurrentDate     =   41023
      MinDate         =   2
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Guardar"
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
      Left            =   360
      Picture         =   "Mov_caja.frx":10DB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Text            =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "L13= nremito"
      Height          =   255
      Left            =   1680
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "L12= nuevbo saldo"
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
      Left            =   240
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "L11= saldo"
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
      Left            =   240
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "L10=id_rrhh"
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
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Listado desde:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Egreso: $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingreso: $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha L2"
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos de caja"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Mov_caja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Select Case Combo1
Case "VIVERO"
    Combo2.Visible = True
    
    Combo2 = "VIVERO SAN NICOLAS SA"
    ' INGRESO falso
    Label5.Visible = False
    Text3.Visible = False
    ' EGRESO verdadero
    Label4.Visible = True
    Text2.Visible = True
Case "Personal"
    Combo2.Clear
    Combo2.Visible = True
    b = "Select * from alta_rrhh order by nombre_rrhh"
    TABLA.Open b, conexion_BD
    Do While Not TABLA.EOF
        Combo2.AddItem TABLA!nombre_rrhh
        TABLA.MoveNext
    Loop
    TABLA.Close
    ' INGRESO false
   Label5.Visible = True
    Text3.Visible = True
    ' EGRESO true
    Label4.Visible = False
    Text2.Visible = False
    
Case "Proveedores"
    Combo2.Visible = True
    Combo2.Clear
    C = "select * from proveedores order by nombre_prove"
    TABLA.Open C, conexion_BD
    Do While Not TABLA.EOF
        Combo2.AddItem TABLA!nombre_prove
    TABLA.MoveNext
    Loop
    TABLA.Close
    ' INGRESO false
   Label5.Visible = True
    Text3.Visible = True
    ' EGRESO true
    Label4.Visible = False
    Text2.Visible = False
    
    
Case "Clientes"
    Combo2.Clear
    Combo2.Visible = True
    
    d = "select * from clientes order by nombre_cliente"
    TABLA.Open d, conexion_BD
    Do While Not TABLA.EOF
        Combo2.AddItem TABLA!nombre_cliente
    TABLA.MoveNext
    Loop
    TABLA.Close
   ' INGRESO false
   Label5.Visible = True
    Text3.Visible = True
    ' EGRESO true
    Label4.Visible = False
    Text2.Visible = False
      
    
End Select
End Sub


Private Sub Combo2_Click()
If Combo1 = "Personal" Then
    j = "Select * from mov_rrhh where nombre_rrhh= '" & Combo2 & "'"
    TABLA.Open j, conexion_BD
   
    If TABLA.EOF = False Then
        K = TABLA!id_rrhh
        'L = TABLA!saldo
        TABLA.Close
    Else
        TABLA.Close
        m = "Select * from alta_rrhh where nombre_rrhh= '" & Combo2 & "'"
        TABLA.Open m, conexion_BD
        K = TABLA!id_rrhh
        'L = 0
        TABLA.Close
        
    End If
    Label10 = K
    'Label11 = L
    
    
End If
End Sub

Private Sub Command1_Click()
If Text1 = "" Or Combo2 = "" Then
    MsgBox "Debe cargar el detalle y el Destinatario", vbOKOnly, "VIVERO SAN NICOLAS"
    Text1.SetFocus
    If Text2 = "" Then
        Text2 = "0"
    
    End If
Else
    
    txtmon = CDbl(Text3)
    Call CONVERTIR
    Label14 = txtmonl
    
    remito = "select max(n_remito) from remito_interno"
    TABLA.Open remito, conexion_BD
    Label13 = TABLA.Fields(0) + 1
    
    Call Guardar
    Call IMPRIMIR
        
End If
Combo2.Clear
Label10 = ""
Label11 = ""
Label12 = ""
Label13 = ""
End Sub

Private Sub Command2_Click()
LISTADO = "select * from mov_caja order by r_interno desc"
TABLA.Open LISTADO, conexion_BD

saldo = 0
saldo1 = 0
resta = 0

    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!detalle
        MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!ingreso, "currency")
        MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!egreso, "currency")
        MSFlexGrid1.TextMatrix(lin, 3) = TABLA!r_interno
        
        saldo = CDbl(saldo) + TABLA!ingreso
        saldo1 = CDbl(saldo1) + TABLA!egreso
        
        TABLA.MoveNext
        
    Loop
TABLA.Close
resta = saldo - saldo1
If resta < 0 Then
    Text4.BackColor = &HFF&
    Text4.ForeColor = &HFFFFFF
    
Else
    Text4.BackColor = vbGreen
    Text4.ForeColor = vbBlack
    
End If
Text4 = Format(resta, "currency")
MSFlexGrid1.Visible = True
Label7.Visible = True

End Sub

Private Sub Command3_Click()
Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)

LISTADO = "select * from mov_caja where fecha >= " + CStr(inicio) + " and fecha <=" + CStr(final) + " order by r_interno desc"
TABLA.Open LISTADO, conexion_BD

MSFlexGrid1.Visible = True
Label7.Visible = True

saldo = 0
saldo1 = 0
resta = 0

    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!detalle
        MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!ingreso, "currency")
        MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!egreso, "currency")
        MSFlexGrid1.TextMatrix(lin, 4) = TABLA!r_interno
        saldo = CDbl(saldo) + TABLA!ingreso
        saldo1 = CDbl(saldo1) + TABLA!egreso
        
        TABLA.MoveNext
    Loop
TABLA.Close


End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()

Label2 = Date
DTPicker1 = Date
DTPicker2 = Date
Call ALTAGRID
End Sub

Private Sub ALTAGRID()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 5
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "FECHA"
MSFlexGrid1.TextMatrix(0, 1) = "DETALLE"
MSFlexGrid1.TextMatrix(0, 2) = "INGRESO"
MSFlexGrid1.TextMatrix(0, 3) = "EGRESO"
MSFlexGrid1.TextMatrix(0, 4) = "REMITO"

MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.ColWidth(4) = 0

End Sub
Private Sub IMPRIMIR()
If Text3 = "" Then
    Text3 = "0"
End If
If Text2 = "" Then
    Text2 = "0"
End If
suma = CDbl(Text3) + CDbl(Text2)


TABLA.Close
rint = "insert into remito_interno values (" & Val(Label13) & ",'" & Label2 & "','" & VIVERO & "','" & Combo2 & "','" & Text1 & "','" & suma & "','" & usua & "')"
conexion_BD.Execute rint

Printer.CurrentX = 30
Printer.CurrentY = 80
Printer.FontSize = 12
Printer.PaperSize = 9 ' papel A4

Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); " Remito Nº: "; Label13
Printer.Print Tab(80); Label2
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí de Vivero San Nicolas S.A. ";
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos: "; suma;
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " ( "; Label14; " ) "
Printer.Print Tab(15); ""
Printer.Print Tab(15); " en concepto de "; Text1; "."
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.CurrentX = 60
Printer.Print Tab(15); " Son  $ "; suma
Printer.Print Tab(60); "_____________";
Printer.Print Tab(58); Combo2.Text
Printer.Print Tab(15); ""
Printer.Print Tab(15); "----------------------------------------------------------------------------------------------------------------";

Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); " Remito Nº: "; Label13
Printer.Print Tab(80); Label2
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí de Vivero San Nicolas S.A. ";
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos: "; suma;
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " ( "; Label14; " ) "
Printer.Print Tab(15); ""
Printer.Print Tab(15); " en concepto de "; Text1; "."
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.CurrentX = 60
Printer.Print Tab(15); " Son  $ "; suma
Printer.Print Tab(60); "_____________";
Printer.Print Tab(58); Combo2.Text
Printer.Print Tab(15); ""
Printer.Print Tab(15); "----------------------------------------------------------------------------------------------------------------";

Printer.EndDoc
End Sub
Private Sub Guardar()

Select Case Combo1.Text
Case "VIVERO"
    a = "insert into mov_caja values ('" & Label2 & "','" & Text1 & "','" & Text2 & "','" & 0 & "','" & Label13 & "')"
    conexion_BD.Execute a
    
Case "Personal"
    Label12 = Label11 - Text3
    'saldo = Val(Label11) - Val(Text3)
    b = "insert into mov_rrhh values (" & Val(Label10) & ",'" & Combo2 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" _
    & 0 & "','" & Text3 & "','" & Text1 & "','" & 0 & "','" & 0 & "','" & Label2 & "','" & 0 & "')"
    conexion_BD.Execute b
    bb = "insert into mov_caja values ('" & Label2 & "','" & Text1 & " " & Combo2 & "','" & 0 & "','" & Text3 & "','" & Label13 & "')"
    conexion_BD.Execute bb

Case "Proveedores"
    
    C = "insert into mov_proveedor values ('" & Combo2 & "','" & Text1 & "','" & Text3 & "','" & 0 & "','" & Label2 & "','" & Label13 & "','EFECTIVO')"
    conexion_BD.Execute C
    cc = "insert into mov_caja values ('" & Label2 & "','" & Text1 & " " & Combo2 & "','" & 0 & "','" & Text3 & "','" & Label13 & "')"
    conexion_BD.Execute cc
    
Case "Clientes"
   
    d = "insert into mov_clientes values ('" & Combo2 & "','" & Text1 & "','" & Text3 & "','" & 0 & "','" & Label2 & "','" & Label13 & "','EFECTIVO')"
    conexion_BD.Execute d
    dd = "insert into mov_caja values ('" & Label2 & "','" & Text1 & " " & Combo2 & "','" & 0 & "','" & Text3 & "','" & Label13 & "')"
    conexion_BD.Execute dd
    
End Select
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

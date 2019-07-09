VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mov_jornales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moviemientos de Jornales "
   ClientHeight    =   7755
   ClientLeft      =   1980
   ClientTop       =   960
   ClientWidth     =   10785
   Icon            =   "Listado_personal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10785
   Begin VB.Frame Frame2 
      Caption         =   "Movimientos realizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   9975
      Begin VB.CommandButton Command3 
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
         Left            =   240
         Picture         =   "Listado_personal.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6480
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   5415
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9551
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3855
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6800
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
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         Picture         =   "Listado_personal.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
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
         Format          =   63045633
         CurrentDate     =   41037
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
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
         Format          =   63045633
         CurrentDate     =   41037
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         Picture         =   "Listado_personal.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5280
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4095
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7223
         _Version        =   393216
         BackColor       =   16777152
         SelectionMode   =   1
         MergeCells      =   1
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
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L5"
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
         Left            =   6960
         TabIndex        =   11
         Top             =   6600
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   5880
         TabIndex        =   10
         Top             =   6600
         Width           =   1095
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
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar desde:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione el nombre que desea buscar."
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
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
      End
   End
End
Attribute VB_Name = "mov_jornales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
d = "select * from mov_rrhh where id_rrhh = " & Val(Combo1) & " order by fecha_mod"
TABLA.Open d, conexion_BD

If TABLA.EOF = False Then
    lin = 0
    MSFlexGrid3.Clear
    Call Alta_personal
    MSFlexGrid2.Visible = False
    MSFlexGrid1.Visible = False
    MSFlexGrid3.Visible = True
    Label5 = ""
    tot = 0
    sueldo = 0
    premio = 0
    adelanto = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
        MSFlexGrid3.TextMatrix(lin, 0) = TABLA!fecha_mod
        MSFlexGrid3.TextMatrix(lin, 1) = TABLA!mes_liquidacion
        MSFlexGrid3.TextMatrix(lin, 2) = TABLA!adelanto
        MSFlexGrid3.TextMatrix(lin, 3) = TABLA!premio
        If TABLA!premio = "" Then
        premio = 0
        Else
        premio = CDbl(premio) + TABLA!premio
        End If
        sueldo = CDbl(sueldo) + TABLA!total_pesos
        adelanto = CDbl(adelanto) + TABLA!adelanto

        TABLA.MoveNext
    Loop
    suma = CDbl(sueldo) + CDbl(premio)
    saldo = CDbl(suma) - CDbl(adelanto)
    Label5 = Format(saldo, "currency")
    
    End If
   TABLA.Close
If saldo < 0 Then
    Label5.BackColor = vbRed
Else
    Label5.BackColor = vbGreen
End If
End Sub
Private Sub USAR_DESPUES()
'COMMAND 1
d = "select * from mov_rrhh where id_rrhh = " & Val(Combo1) & " order by fecha_mod"
TABLA.Open d, conexion_BD

If TABLA.EOF = False Then
    lin = 0
    MSFlexGrid2.Clear
    Call Alta_personal
    MSFlexGrid2.Visible = True
    MSFlexGrid1.Visible = False
    Label5 = ""
    tot = 0
    sueldo = 0
    premio = 0
    adelanto = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!fecha_mod
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!mes_liquidacion
        MSFlexGrid2.TextMatrix(lin, 2) = TABLA!hs_normal
        MSFlexGrid2.TextMatrix(lin, 3) = TABLA!hs_50
        MSFlexGrid2.TextMatrix(lin, 4) = TABLA!hs_100
        MSFlexGrid2.TextMatrix(lin, 5) = TABLA!total_hs
        MSFlexGrid2.TextMatrix(lin, 6) = TABLA!total_pesos
        MSFlexGrid2.TextMatrix(lin, 7) = TABLA!adelanto
        MSFlexGrid2.TextMatrix(lin, 8) = TABLA!precioxhora
        MSFlexGrid2.TextMatrix(lin, 9) = TABLA!premio
        If TABLA!premio = "" Then
        premio = 0
        Else
        premio = CDbl(premio) + TABLA!premio
        End If
        sueldo = CDbl(sueldo) + TABLA!total_pesos
        adelanto = CDbl(adelanto) + TABLA!adelanto

        TABLA.MoveNext
    Loop
    suma = CDbl(sueldo) + CDbl(premio)
    saldo = CDbl(suma) - CDbl(adelanto)
    Label5 = Format(saldo, "currency")
    
    End If
   TABLA.Close
If saldo < 0 Then
    Label5.BackColor = vbRed
Else
    Label5.BackColor = vbGreen
End If


' COMMAND 2
'e = "select * from mov_rrhh where fecha_mod between # " & DTPicker1 & " #  and # " & DTPicker2 & " # "
'TABLA.Open e, conexion_BD

e = "select * from mov_rrhh where fecha_mod >= # " & DTPicker1 & " #  and fecha_mod <= # " & DTPicker2 & " # "
TABLA.Open e, conexion_BD

MSFlexGrid2.Visible = False
MSFlexGrid1.Visible = True
Label4.Visible = False
Label5.Visible = False
Call alta1
If TABLA.EOF = False Then
    With MSFlexGrid1
    lin = 0
    
    .Visible = True
    Label5 = ""
    tot = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        .TextMatrix(lin, 0) = TABLA!nombre_rrhh
        .TextMatrix(lin, 1) = TABLA!fecha_mod
        .TextMatrix(lin, 2) = TABLA!mes_liquidacion
        .TextMatrix(lin, 3) = TABLA!hs_normal
        .TextMatrix(lin, 4) = TABLA!hs_50
        .TextMatrix(lin, 5) = TABLA!hs_100
        .TextMatrix(lin, 6) = TABLA!total_hs
        .TextMatrix(lin, 7) = TABLA!total_pesos
        .TextMatrix(lin, 8) = TABLA!adelanto
        .TextMatrix(lin, 9) = TABLA!premio
       
        If TABLA!premio = "" Then
        premio = 0
        Else
        premio = CDbl(premio) + TABLA!premio
        End If
        sueldo = CDbl(sueldo) + TABLA!total_pesos
        adelanto = CDbl(adelanto) + TABLA!adelanto
        TABLA.MoveNext
    Loop
    End With
    suma = CDbl(sueldo) + CDbl(premio)
    saldo = CDbl(suma) - CDbl(adelanto)
    Label5 = Format(saldo, "currency")
    End If
   TABLA.Close
   
' LOAD FORM

DTPicker1 = Date
DTPicker2 = Date
Label5 = ""
Frame2.Visible = True
Call Alta_personal
If datos = "" Then
    co = "select * from alta_rrhh order by id_rrhh"
    TABLA.Open co, conexion_BD
    Do While Not TABLA.EOF
        Combo1.AddItem TABLA!id_rrhh & " " & TABLA!nombre_rrhh
        TABLA.MoveNext
    Loop
    TABLA.Close

Else
    e = "select * from mov_rrhh where id_rrhh = " & Val(datos) & " " 'order by fecha_mod"
    TABLA.Open e, conexion_BD

    MSFlexGrid2.Visible = False
    MSFlexGrid1.Visible = True

    Call alta1
    If TABLA.EOF = False Then
        With MSFlexGrid1
        lin = 0
    
        .Visible = True
        Label5 = ""
        tot = 0
        Do While Not TABLA.EOF
            lin = lin + 1
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            .TextMatrix(lin, 0) = TABLA!nombre_rrhh
            .TextMatrix(lin, 1) = TABLA!fecha_mod
            .TextMatrix(lin, 2) = TABLA!mes_liquidacion
            .TextMatrix(lin, 3) = TABLA!hs_normal
            .TextMatrix(lin, 4) = TABLA!hs_50
            .TextMatrix(lin, 5) = TABLA!hs_100
            .TextMatrix(lin, 6) = TABLA!total_hs
            .TextMatrix(lin, 7) = TABLA!total_pesos
            .TextMatrix(lin, 8) = TABLA!adelanto
            .TextMatrix(lin, 9) = TABLA!precioxhora
            .TextMatrix(lin, 10) = TABLA!premio
       
            If TABLA!premio = "" Then
                premio = 0
            Else
                premio = CDbl(premio) + TABLA!premio
            End If
        
            sueldo = CDbl(sueldo) + TABLA!total_pesos
            adelanto = CDbl(adelanto) + TABLA!adelanto
            TABLA.MoveNext
        Loop
        End With
        suma = CDbl(sueldo) + CDbl(premio)
        saldo = CDbl(suma) - CDbl(adelanto)
        Label5 = Format(saldo, "currency")
    
        End If
    TABLA.Close
End If
If saldo < 0 Then
    Label5.BackColor = vbRed
Else
    Label5.BackColor = vbGreen
End If
End Sub

Private Sub Command2_Click()

e = "select * from mov_rrhh where fecha_mod >= # " & DTPicker1 & " #  and fecha_mod <= # " & DTPicker2 & " # "
TABLA.Open e, conexion_BD

MSFlexGrid2.Visible = False
MSFlexGrid1.Visible = False
MSFlexGrid3.Visible = True
Label4.Visible = False
Label5.Visible = False
Call alta1
If TABLA.EOF = False Then
    With MSFlexGrid3
    lin = 0
    
    .Visible = True
    Label5 = ""
    tot = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!nombre_rrhh
        .TextMatrix(lin, 1) = TABLA!fecha_mod
        .TextMatrix(lin, 2) = TABLA!mes_liquidacion
        .TextMatrix(lin, 3) = TABLA!total_pesos
        .TextMatrix(lin, 4) = TABLA!adelanto
        .TextMatrix(lin, 5) = TABLA!premio
       
        If TABLA!premio = "" Then
        premio = 0
        Else
        premio = CDbl(premio) + TABLA!premio
        End If
        sueldo = CDbl(sueldo) + TABLA!total_pesos
        adelanto = CDbl(adelanto) + TABLA!adelanto
        TABLA.MoveNext
    Loop
    End With
    suma = CDbl(sueldo) + CDbl(premio)
    saldo = CDbl(suma) - CDbl(adelanto)
    Label5 = Format(saldo, "currency")
    End If
   TABLA.Close
End Sub

Private Sub Command3_Click()
Call IMPRIMIR
End Sub
Private Sub IMPRIMIR()
Printer.Orientation = 1
Printer.FontSize = 10
Printer.Font = arial
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Fecha: "; Date; Tab(59); " VIVERO SAN NICOLAS SA";
Printer.Print Tab(50); "RUTA PROVINCIAL 60 S/N, JUNIN, MENDOZA"
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "MOVIMIENTO DEL PERSONAL"
Printer.FontSize = 8
Printer.Print Tab(15); ""
Printer.Print Tab(10); Label1
Printer.Print Tab(15); ""

Printer.Print Tab(10); "FECHA"; Tab(25); "DETALLE"; Tab(65); "MES"; Tab(85); "TOTAL EN $"; Tab(100); "ADELANTOS"; Tab(115); "PREMIO";  'Tab(137); "FECHA MODIF"
Printer.Print Tab(10); "=============================================================================================="

For i = 1 To MSFlexGrid3.Rows - 1
    MSFlexGrid3.Row = i
    fecha = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 0)
    detalle = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 1)
    mes = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 2)
    total = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 3)
    adelantos = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 4)
    premio = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 5)


Printer.Print Tab(10); fecha; Tab(25); detalle; Tab(65); mes; Tab(87); total; Tab(102); adelantos; Tab(117); premio; ' Tab(139); depo
Printer.Print Tab(10); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------"

If i = 36 Then
    Printer.NewPage
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); "FECHA"; Tab(25); "DETALLE"; Tab(65); "MES"; Tab(85); "TOTAL EN $"; Tab(100); "ADELANTOS"; Tab(115); "PREMIO";  'Tab(137); "FECHA MODIF"
    Printer.Print Tab(10); "=============================================================================================="

End If

Next
Printer.Print Tab(10); ""
Printer.Print Tab(10); " TOTAL: "; Label5
Printer.EndDoc
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
Label5 = ""
Frame2.Visible = True
'Call Alta_personal
If datos = "" Then
    co = "select * from alta_rrhh order by id_rrhh"
    TABLA.Open co, conexion_BD
    Do While Not TABLA.EOF
        Combo1.AddItem TABLA!id_rrhh & " " & TABLA!nombre_rrhh
        TABLA.MoveNext
    Loop
    TABLA.Close

Else
    'e = "select * from mov_rrhh where id_rrhh = " & Val(datos) & " order by fecha_mod"
    'TABLA.Open e, conexion_BD

    e = "select * from mov_rrhh where nombre_rrhh = '" & datos1 & "' order by fecha_mod"
    TABLA.Open e, conexion_BD

    MSFlexGrid2.Visible = False
    MSFlexGrid1.Visible = False
    MSFlexGrid3.Visible = True
    
    Label1 = datos1
    Label1.Visible = True

    Call alta1
    If TABLA.EOF = False Then
        With MSFlexGrid3
        lin = 0
    
        .Visible = True
        Label5 = ""
        tot = 0
        premio = 0
        Do While Not TABLA.EOF
            lin = lin + 1
            MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
            .TextMatrix(lin, 0) = TABLA!fecha_mod
            .TextMatrix(lin, 1) = TABLA!detalle
            .TextMatrix(lin, 2) = TABLA!mes_liquidacion
            .TextMatrix(lin, 3) = TABLA!total_pesos
            .TextMatrix(lin, 4) = TABLA!adelanto
            .TextMatrix(lin, 5) = TABLA!premio
       
            If TABLA!premio = "" Then
                premio = 0 + CDbl(premio)
            Else
                premio = CDbl(premio) + TABLA!premio
            End If
        
            sueldo = CDbl(sueldo) + TABLA!total_pesos
            adelanto = CDbl(adelanto) + TABLA!adelanto
            TABLA.MoveNext
        Loop
        End With
        suma = CDbl(sueldo) + CDbl(premio)
        saldo = CDbl(suma) - CDbl(adelanto)
        Label5 = Format(saldo, "currency")
    
        End If
    TABLA.Close
End If
If saldo < 0 Then
    Label5.BackColor = vbRed
Else
    Label5.BackColor = vbGreen
End If
End Sub
Private Sub Alta_personal()
MSFlexGrid3.FixedCols = 0
MSFlexGrid3.Cols = 5
MSFlexGrid3.FixedRows = 1
MSFlexGrid3.Rows = 2
MSFlexGrid3.Clear
MSFlexGrid3.TextMatrix(0, 0) = "FECHA"
MSFlexGrid3.TextMatrix(0, 1) = "MES"
MSFlexGrid3.TextMatrix(0, 2) = "TOTAL EN $"
MSFlexGrid3.TextMatrix(0, 3) = "ADELANTOS"
MSFlexGrid3.TextMatrix(0, 4) = "PREMIO"

MSFlexGrid3.ColWidth(0) = 2500
MSFlexGrid3.ColWidth(1) = 2000
MSFlexGrid3.ColWidth(2) = 1000
MSFlexGrid3.ColWidth(3) = 1000
MSFlexGrid3.ColWidth(4) = 1000



End Sub
Private Sub alta1()

MSFlexGrid3.FixedCols = 0
MSFlexGrid3.Cols = 6
MSFlexGrid3.FixedRows = 1
MSFlexGrid3.Rows = 2
MSFlexGrid3.Clear
MSFlexGrid3.TextMatrix(0, 0) = "FECHA"
MSFlexGrid3.TextMatrix(0, 1) = "DETALLE"
MSFlexGrid3.TextMatrix(0, 2) = "MES"
MSFlexGrid3.TextMatrix(0, 3) = "TOTAL EN $"
MSFlexGrid3.TextMatrix(0, 4) = "ADELANTOS"
MSFlexGrid3.TextMatrix(0, 5) = "PREMIO"


MSFlexGrid3.ColWidth(0) = 1500
MSFlexGrid3.ColWidth(1) = 3000
MSFlexGrid3.ColWidth(2) = 1000
MSFlexGrid3.ColWidth(3) = 1000
MSFlexGrid3.ColWidth(4) = 1000
MSFlexGrid3.ColWidth(5) = 1000



End Sub


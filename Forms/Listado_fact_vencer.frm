VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Listado_fact_vencer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Facturas a Vencer"
   ClientHeight    =   9360
   ClientLeft      =   720
   ClientTop       =   915
   ClientWidth     =   13785
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Listado_fact_vencer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13785
   Begin VB.CommandButton Command1 
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
      Left            =   11520
      Picture         =   "Listado_fact_vencer.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Height          =   6375
      Left            =   2760
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   6975
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   5655
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9975
         _Version        =   393216
         BackColor       =   16777152
         SelectionMode   =   1
      End
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
      Left            =   9120
      Picture         =   "Listado_fact_vencer.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
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
      Left            =   10320
      Picture         =   "Listado_fact_vencer.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtrar por Fecha"
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   5055
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   360
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
         Format          =   61800449
         CurrentDate     =   40909
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   360
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
         Format          =   61800449
         CurrentDate     =   41082
      End
      Begin VB.Label Label3 
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
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Desde:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar por:"
      Height          =   855
      Left            =   5520
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
      Begin VB.OptionButton Option2 
         Caption         =   "Proveedores"
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
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clientes"
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
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6135
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   12975
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5655
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   9975
         _Version        =   393216
         BackColor       =   16777152
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame4 
      Height          =   6375
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   13215
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   5895
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   10398
         _Version        =   393216
         BackColor       =   16777152
         SelectionMode   =   1
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Listados de Facturas a Vencer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "Listado_fact_vencer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub IMPRIMIR()
Printer.FontSize = 12
Printer.Font = arial
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Fecha: "; Date; Tab(55); SISTEMA; Tab(90); Page
Printer.Print Tab(50); SISTEMA_DIR
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(30); "LISTADO DE CUENTAS CON FACTURAS A VENCER"
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.FontSize = 10
Printer.Print Tab(15); " Detalle desde "; DTPicker1; " Hasta "; DTPicker2
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""

For i = 0 To MSFlexGrid3.Rows - 1
    MSFlexGrid3.Row = i
    nombre = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 0)
    monto = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 1)
    Printer.FontSize = 6
    Printer.Print Tab(15); ""
    Printer.FontSize = 10
    
Printer.Print Tab(15); nombre; Tab(85); monto;

If i = 33 Then
    Printer.NewPage
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""

End If

Next

Printer.EndDoc

End Sub
Private Sub Command1_Click()
Call IMPRIMIR
Command1.Visible = False
End Sub

Private Sub Command2_Click()
    Frame3.Visible = False
    Frame4.Visible = False
    Frame5.Visible = True
    Frame5.BorderStyle = 0
    Command1.Visible = True
    
    inicio = CDate(Me.DTPicker1.value)
    final = CDate(Me.DTPicker2.value)
    
If Option1 = True Then

    LISTADO = "select * from clientes_a_pagar where fecha_vto >= #" & DTPicker1 & "# and fecha_vto <= #" & DTPicker2 & "# order by cliente"
    TABLA.Open LISTADO, conexion_BD
    
    Call ALTA_FECHAS
    lin = 0

    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
        
        If TABLA!verdadero = "-1" Then '''' SI ESTA AFECTADA LA FACTURA ELIMINA EL REGISTRO
            MSFlexGrid3.Row = lin
            MSFlexGrid3.RemoveItem (MSFlexGrid3.Row)
            lin = lin - 1
        Else
            ''' SI NO ESTA AFECTADA LA FACTURA  TRAE EL REGISTRO
            MSFlexGrid3.TextMatrix(lin, 0) = TABLA!cliente
            tot = CDbl(tot) + TABLA!monto
            'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
            If MSFlexGrid3.TextMatrix(lin, 0) = MSFlexGrid3.TextMatrix(lin - 1, 0) Then
                
                MSFlexGrid3.TextMatrix(lin, 1) = TABLA!monto
                
                cta = CDbl(MSFlexGrid3.TextMatrix(lin, 1)) + CDbl(MSFlexGrid3.TextMatrix(lin - 1, 1))
                
                MSFlexGrid3.Row = lin
                MSFlexGrid3.RemoveItem (MSFlexGrid3.Row)
                
                MSFlexGrid3.TextMatrix(lin - 1, 1) = cta
                lin = lin - 1
            Else
            
                MSFlexGrid3.TextMatrix(lin, 0) = TABLA!cliente
                MSFlexGrid3.TextMatrix(lin, 1) = TABLA!monto
        
            End If

        End If
        TABLA.MoveNext
    Loop
    TABLA.Close
Else
    LISTADO = "select * from prove_a_pagar where fecha_vto >= #" & DTPicker1 & "# and fecha_vto <= #" & DTPicker2 & "# order by proveedor"
    TABLA.Open LISTADO, conexion_BD
    Call ALTA_FECHAS
    lin = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        
        MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
        
        If TABLA!verdadero = "-1" Then '''' SI ESTA AFECTADA LA FACTURA ELIMINA EL REGISTRO
            MSFlexGrid3.Row = lin
            MSFlexGrid3.RemoveItem (MSFlexGrid3.Row)
            lin = lin - 1
        Else
            ''' SI NO ESTA AFECTADA LA FACTURA  TRAE EL REGISTRO
            MSFlexGrid3.TextMatrix(lin, 0) = TABLA!proveedor
            If MSFlexGrid3.TextMatrix(lin, 0) = MSFlexGrid3.TextMatrix(lin - 1, 0) Then
            
                MSFlexGrid3.TextMatrix(lin, 1) = TABLA!monto
                
                cta = CDbl(MSFlexGrid3.TextMatrix(lin, 1)) + CDbl(MSFlexGrid3.TextMatrix(lin - 1, 1))
                
                MSFlexGrid3.Row = lin
                MSFlexGrid3.RemoveItem (MSFlexGrid3.Row)
                
                MSFlexGrid3.TextMatrix(lin - 1, 1) = cta
                lin = lin - 1
 
            Else
            
                MSFlexGrid3.TextMatrix(lin, 0) = TABLA!proveedor
                MSFlexGrid3.TextMatrix(lin, 1) = TABLA!monto
                cuenta = 0
            End If
            
   
        End If
        TABLA.MoveNext
    Loop
    TABLA.Close
End If

End Sub

Private Sub Command3_Click()
If Option1 = True Then

    Call Altas_Clientes
    Frame3.BorderStyle = 0
    Frame3.Visible = True
    Frame4.Visible = False
    Frame5.Visible = False

    Dim inicio As Long
    Dim final As Long

    inicio = CDate(Me.DTPicker1.value)
    final = CDate(Me.DTPicker2.value)

    LISTADO = "select * from clientes_a_pagar where fecha_vto >= #" & DTPicker1 & "# and fecha_vto <= #" & DTPicker2 & "# order by fecha_vto desc"
    TABLA.Open LISTADO, conexion_BD
    
    lin = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!fecha_vto
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!cliente
        MSFlexGrid1.TextMatrix(lin, 2) = TABLA!factura
        MSFlexGrid1.TextMatrix(lin, 3) = TABLA!detalle
        MSFlexGrid1.TextMatrix(lin, 4) = TABLA!monto
        MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha
        If TABLA!verdadero = "-1" Then
            MSFlexGrid1.Row = lin
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
            lin = lin - 1
        End If

        TABLA.MoveNext
    Loop
    
    TABLA.Close
Else

    Call Altas_Proveedores
    Frame4.BorderStyle = 0
    Frame4.Visible = True
    Frame3.Visible = False
    Frame5.Visible = False

    'listado = "select * from prove_a_pagar where fecha_vto between # " & DTPicker1 & "# and # " & DTPicker2 & "# order by fecha_vto"
    'TABLA.Open listado, conexion_BD
    
    LISTADO = "select * from prove_a_pagar where fecha_vto >= #" & DTPicker1 & "# and fecha_vto <= #" & DTPicker2 & "# order by fecha_vto"
    TABLA.Open LISTADO, conexion_BD
    
    lin = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!fecha_vto
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!proveedor
        MSFlexGrid2.TextMatrix(lin, 2) = TABLA!factura
        MSFlexGrid2.TextMatrix(lin, 3) = TABLA!detalle
        MSFlexGrid2.TextMatrix(lin, 4) = TABLA!monto
        MSFlexGrid2.TextMatrix(lin, 5) = TABLA!fecha
        If TABLA!verdadero = "-1" Then
            MSFlexGrid2.Row = lin
            MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
            lin = lin - 1
        End If
        TABLA.MoveNext
    Loop
    
    TABLA.Close
End If
End Sub

Private Sub Form_Load()

DTPicker2 = Date
End Sub
Private Sub Altas_Proveedores()
MSFlexGrid2.FixedCols = 0
MSFlexGrid2.Cols = 6
MSFlexGrid2.FixedRows = 1
MSFlexGrid2.Rows = 2
MSFlexGrid2.Clear
MSFlexGrid2.TextMatrix(0, 0) = "FECHA VTO"
MSFlexGrid2.TextMatrix(0, 1) = "CLIENTE"
MSFlexGrid2.TextMatrix(0, 2) = "FACTURA"
MSFlexGrid2.TextMatrix(0, 3) = "DETALLE"
MSFlexGrid2.TextMatrix(0, 4) = "MONTO"
MSFlexGrid2.TextMatrix(0, 5) = "FECHA INGRESO"


MSFlexGrid2.ColWidth(0) = 1500
MSFlexGrid2.ColWidth(1) = 3000
MSFlexGrid2.ColWidth(2) = 1500
MSFlexGrid2.ColWidth(3) = 3000
MSFlexGrid2.ColWidth(4) = 1500
MSFlexGrid2.ColWidth(5) = 1500
End Sub

Private Sub Altas_Clientes()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 6
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "FECHA VTO"
MSFlexGrid1.TextMatrix(0, 1) = "CLIENTE"
MSFlexGrid1.TextMatrix(0, 2) = "FACTURA"
MSFlexGrid1.TextMatrix(0, 3) = "DETALLE"
MSFlexGrid1.TextMatrix(0, 4) = "MONTO"
MSFlexGrid1.TextMatrix(0, 5) = "FECHA INGRESO"


MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 3000
MSFlexGrid1.ColWidth(4) = 1500
MSFlexGrid1.ColWidth(5) = 1500
End Sub
Private Sub ALTA_FECHAS()

MSFlexGrid3.FixedCols = 0
MSFlexGrid3.Cols = 2
MSFlexGrid3.FixedRows = 1
MSFlexGrid3.Rows = 2
MSFlexGrid3.Clear
MSFlexGrid3.TextMatrix(0, 0) = "NOMBRE/RAZON SOCIAL"
MSFlexGrid3.TextMatrix(0, 1) = "MONTO"

MSFlexGrid3.ColWidth(0) = 4200
MSFlexGrid3.ColWidth(1) = 1500

End Sub


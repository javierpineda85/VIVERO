VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Afectar_factura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Afectación de factura"
   ClientHeight    =   9360
   ClientLeft      =   1095
   ClientTop       =   885
   ClientWidth     =   13800
   Icon            =   "Afectar_factura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13800
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   8280
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancelar"
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
         Left            =   1920
         Picture         =   "Afectar_factura.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Continuar"
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
         Left            =   720
         Picture         =   "Afectar_factura.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Factura Nº:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Monto a pagar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha vto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   2880
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos ocultos"
      Height          =   1335
      Left            =   9840
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "L8=fecha_vto"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "L7=fecha_fact"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "L6=fecha"
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "L5=monto"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "L4=detalle"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "L3=factura"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "L2= cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7095
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12515
      _Version        =   393216
      BackColor       =   16777152
      BackColorBkg    =   -2147483636
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
   Begin VB.Label Label1 
      Caption         =   "Afectar factura a un pago"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   600
      Width           =   5295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Afectar_factura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 
saldo1 = CDbl(Label17) - CDbl(Text1)
datos1 = Label15
datos2 = Text1
datos3 = saldo1
datos4 = Label16
datos8 = Label5 'monto de la factura
tot = CDbl(tot) + CDbl(Text1)
datos6 = CDbl(tot) + CDbl(datos6)
    
Frame2.Visible = False
MSFlexGrid1.RemoveItem (MSFlexGrid1.RowSel)
Label19 = Format(datos6, "currency")
End Sub

Private Sub Command2_Click()
Frame2.Visible = False
End Sub

Private Sub Form_Load()
Call INI
tot = 0
datos6 = 0
End Sub

Private Sub Alta()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 7
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "Nº FACTURA"
MSFlexGrid1.TextMatrix(0, 1) = "DETALLE"
MSFlexGrid1.TextMatrix(0, 2) = "MONTO"
MSFlexGrid1.TextMatrix(0, 3) = "FECHA PROCESO"
MSFlexGrid1.TextMatrix(0, 4) = "FECHA FACTURA"
MSFlexGrid1.TextMatrix(0, 5) = "FECHA VENC."


MSFlexGrid1.ColWidth(0) = 2000
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 2000
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 2000
MSFlexGrid1.ColWidth(5) = 2000


End Sub

Private Sub MSFlexGrid1_Click()
Label2 = datos
Label3 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
Label5 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
Label9 = "-1"
datos5 = Label3 ' con llevamos el nº de factura a los proveedores o clientes

preg = MsgBox("Desea realizar un pago total?", vbYesNo)

If preg = vbYes Then
    datos7 = "TOTAL" ' llevamos con esta variable a proveedores y cliente la info si es parcial o total

''' PAGOS TOTALES '''
    
    '''''''TOMAMOS LOS DATOS PARA GUARDAR EN TABLA FACTURAS"
    factura = Label3
    cliente = datos
    detalle = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    monto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    fecha = Date
    fecha_vto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 5)
    saldo = 0
    estado = "CANCELADO"
    pparcial = 0
    r_interno = 0

    If sValor = "CLIENTES" Then
        resp = MsgBox(" Desea afectar esta factura al pago? IMPORTANTE: Una vez afectado no se puede deshacer el movimiento", vbOKCancel)
            If resp = vbOK Then
                datos6 = CDbl(datos6) + Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
                L = "update clientes_a_pagar set verdadero=" & Val(Label9) & " where factura ='" & Label3 & "'"
                conexion_BD.Execute L
                
                fra = "insert into facturas values ('" & factura & "','" & cliente & "','" & detalle & "','" & _
                monto & "','" & fecha & "','" & fecha_vto & "','" & saldo & "','" & estado & "','" & _
                pparcial & "','" & r_interno & "')"
                conexion_BD.Execute fra
                
                MSFlexGrid1.RemoveItem (MSFlexGrid1.RowSel)
            End If
    
    Else
        
        resp = MsgBox(" Desea afectar esta factura al pago? IMPORTANTE: Una vez afectado no se puede deshacer el movimiento", vbOKCancel)
            If resp = vbOK Then
                datos6 = CDbl(datos6) + Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
                L = "update prove_a_pagar set verdadero=" & Val(Label9) & " where factura ='" & Label3 & "'"
                conexion_BD.Execute L
                
                fra = "insert into facturas values ('" & factura & "','" & cliente & "','" & detalle & "','" & _
                monto & "','" & fecha & "','" & fecha_vto & "','" & saldo & "','" & estado & "','" & _
                pparcial & "','" & r_interno & "')"
                conexion_BD.Execute fra
                
                MSFlexGrid1.RemoveItem (MSFlexGrid1.RowSel)
            End If
    End If
Else

    datos7 = "PARCIAL"
''' PAGOS PARCIALES '''
''' NO REALIZA LA ACCION. SE REGISTRA EL PAGO PARCIAL DESDE EL MOV_CLIENTES/PROVEEDORE

    
    Frame2.Visible = True
    buscar = " select * from facturas where factura ='" & Label3 & "'"
    TABLA.Open buscar, conexion_BD
    
    factura = Label3
    cliente = datos
    detalle = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    monto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    fecha = Date
    fecha_vto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 5)
    monto = 0
    pparcial = 0
    
    Do While Not TABLA.EOF
    
        monto = TABLA!monto
        pparcial = CDbl(pparcial) + TABLA!pparcial
  
        factura = TABLA!factura
        fecha_vto = TABLA!fecha_vto
        estado = TABLA!estado

    TABLA.MoveNext
    
    Loop
    
    saldo = CDbl(monto) - CDbl(pparcial)
    Label15 = factura
    Label16 = fecha_vto
    Label18 = estado
    Label17 = Format(saldo, "currency")

    TABLA.Close
    

    
    'MSFlexGrid1.RemoveItem (MSFlexGrid1.RowSel)
End If
Label19 = Format(datos6, "currency")
End Sub

Public Sub INI()


If sValor = "CLIENTES" Then

    com = "select * from clientes_a_pagar where cliente ='" & datos & "' order by fecha_vto"
    TABLA.Open com, conexion_BD

    
Else

    com = "select * from prove_a_pagar where proveedor='" & datos & "' order by fecha_vto"
    TABLA.Open com, conexion_BD
    
End If
Label9 = "-1"
Call Alta
lin = 1
Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!factura
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!detalle
        MSFlexGrid1.TextMatrix(lin, 2) = TABLA!monto
        MSFlexGrid1.TextMatrix(lin, 3) = TABLA!fecha
        MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_factura
        MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
        
        If TABLA!verdadero = -1 Then
            MSFlexGrid1.Row = lin
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row + 1)
            lin = lin - 1
        End If
        TABLA.MoveNext
    Loop
    
    TABLA.Close
If lin = 0 Then
    MSFlexGrid1.Clear
    MsgBox " No hay facturas pendientes de pago"
    'Unload Me
End If
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Listado_clientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de clientes"
   ClientHeight    =   9360
   ClientLeft      =   720
   ClientTop       =   855
   ClientWidth     =   13470
   Icon            =   "Listado_clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13470
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6615
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   11668
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Clientes"
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
      Left            =   4680
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
End
Attribute VB_Name = "Listado_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
L = "select * from clientes order by nombre_cliente"
TABLA.Open L, conexion_BD
Call Alta_Cliente
    Do While Not TABLA.EOF
        fil = fil + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(fil, 0) = TABLA!id_cliente
        MSFlexGrid1.TextMatrix(fil, 1) = TABLA!nombre_cliente
        MSFlexGrid1.TextMatrix(fil, 2) = TABLA!CUIT
        MSFlexGrid1.TextMatrix(fil, 3) = TABLA!iva
        MSFlexGrid1.TextMatrix(fil, 4) = TABLA!CTACTE
        MSFlexGrid1.TextMatrix(fil, 5) = TABLA!direccion
        MSFlexGrid1.TextMatrix(fil, 6) = TABLA!localidad
        MSFlexGrid1.TextMatrix(fil, 7) = TABLA!telefono
        MSFlexGrid1.TextMatrix(fil, 8) = TABLA!email
        'MSFlexGrid1.TextMatrix(fil, 9) = TABLA!id_empresa
        TABLA.MoveNext
    Loop
    
    TABLA.Close
    
End Sub

Private Sub Alta_Cliente()

With MSFlexGrid1
.FixedCols = 0
.Cols = 10
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "Nº de PROD."
.TextMatrix(0, 1) = "APELLIDO Y NOMBRE"
.TextMatrix(0, 2) = "CUIT"
.TextMatrix(0, 3) = "IVA"
.TextMatrix(0, 4) = "CTA CTE"
.TextMatrix(0, 5) = "DIRECCION"
.TextMatrix(0, 6) = "LOCALIDAD"
.TextMatrix(0, 7) = "TELEFONO"
.TextMatrix(0, 8) = "CORREO ELECTRONICO"
.TextMatrix(0, 9) = "EMPRESA Nº"

.ColWidth(0) = 800
.ColWidth(1) = 3000
.ColWidth(2) = 2000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 4500
.ColWidth(6) = 2000
.ColWidth(7) = 2000
.ColWidth(8) = 2500
.ColWidth(9) = 800
End With
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Listado_proveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de proveedores"
   ClientHeight    =   9360
   ClientLeft      =   825
   ClientTop       =   915
   ClientWidth     =   13470
   Icon            =   "Listado_proveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13470
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   7215
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12726
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "De Proveedores"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Listado General"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Listado_proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Listado_proveedores.PrintForm
End Sub

Private Sub Form_Load()

com = "select * from proveedores order by nombre_prove asc"
TABLA.Open com, conexion_BD
Call Alta_Proveedores

Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
        MSFlexGrid3.TextMatrix(lin, 0) = TABLA!nombre_prove
        MSFlexGrid3.TextMatrix(lin, 1) = TABLA!direccion
        MSFlexGrid3.TextMatrix(lin, 2) = TABLA!ciudad
        MSFlexGrid3.TextMatrix(lin, 3) = TABLA!provincia
        MSFlexGrid3.TextMatrix(lin, 4) = TABLA!telefono
        MSFlexGrid3.TextMatrix(lin, 5) = TABLA!email
        TABLA.MoveNext
    Loop
    
    TABLA.Close

End Sub

Private Sub Alta_Proveedores()
MSFlexGrid3.FixedCols = 0
MSFlexGrid3.Cols = 6
MSFlexGrid3.FixedRows = 1
MSFlexGrid3.Rows = 2
MSFlexGrid3.Clear
MSFlexGrid3.TextMatrix(0, 0) = "PROVEEDOR"
MSFlexGrid3.TextMatrix(0, 1) = "DIRECCION"
MSFlexGrid3.TextMatrix(0, 2) = "LOCALIDAD"
MSFlexGrid3.TextMatrix(0, 3) = "PROVINCIA"
MSFlexGrid3.TextMatrix(0, 4) = "TELEFONO"
MSFlexGrid3.TextMatrix(0, 5) = "CORREO ELECTRONICO"

MSFlexGrid3.ColWidth(0) = 2500
MSFlexGrid3.ColWidth(1) = 3000
MSFlexGrid3.ColWidth(2) = 2000
MSFlexGrid3.ColWidth(3) = 2000
MSFlexGrid3.ColWidth(4) = 2000
MSFlexGrid3.ColWidth(5) = 4500

End Sub



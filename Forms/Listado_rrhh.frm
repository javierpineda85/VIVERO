VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Listado_rrhh 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Personal"
   ClientHeight    =   9360
   ClientLeft      =   825
   ClientTop       =   960
   ClientWidth     =   13470
   Icon            =   "Listado_rrhh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13470
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7950
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   14023
      _Version        =   393216
      BackColor       =   16777152
      SelectionMode   =   1
      AllowUserResizing=   2
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
      Caption         =   "Listado del Personal"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "Listado_rrhh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
W = "select * from alta_rrhh order by nombre_rrhh asc"
TABLA.Open W, conexion_BD

Call alta_rrhh
    Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!id_rrhh
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!nombre_rrhh
        MSFlexGrid1.TextMatrix(lin, 2) = TABLA!cuil
        MSFlexGrid1.TextMatrix(lin, 3) = TABLA!direccion
        MSFlexGrid1.TextMatrix(lin, 4) = TABLA!ciudad
        MSFlexGrid1.TextMatrix(lin, 5) = TABLA!provincia
        MSFlexGrid1.TextMatrix(lin, 6) = TABLA!telefono
        TABLA.MoveNext
    Loop
    
    TABLA.Close
    
End Sub
Private Sub alta_rrhh()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 7
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "LEGAJO Nº"
MSFlexGrid1.TextMatrix(0, 1) = "APELLIDO Y NOMBRE"
MSFlexGrid1.TextMatrix(0, 2) = "CUIL"
MSFlexGrid1.TextMatrix(0, 3) = "DIRECCION"
MSFlexGrid1.TextMatrix(0, 4) = "LOCALIDAD"
MSFlexGrid1.TextMatrix(0, 5) = "PROVINCIA"
MSFlexGrid1.TextMatrix(0, 6) = "TELEFONO"

MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 2000
MSFlexGrid1.ColWidth(3) = 3000
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1000

End Sub



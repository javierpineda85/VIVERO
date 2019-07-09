VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form alta_proveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Altas de Proveedres"
   ClientHeight    =   9105
   ClientLeft      =   2880
   ClientTop       =   435
   ClientWidth     =   9510
   Icon            =   "alta_proveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   9510
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
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
      Picture         =   "alta_proveedor.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      DisabledPicture =   "alta_proveedor.frx":0B14
      DragIcon        =   "alta_proveedor.frx":109E
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
      Picture         =   "alta_proveedor.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Borrar"
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
      Picture         =   "alta_proveedor.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3255
      Left            =   360
      TabIndex        =   16
      Top             =   5520
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5741
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Editar"
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
      Left            =   4320
      Picture         =   "alta_proveedor.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
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
      Picture         =   "alta_proveedor.frx":26C6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text5 
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
      Left            =   5760
      TabIndex        =   6
      Top             =   3720
      Width           =   3255
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
      ItemData        =   "alta_proveedor.frx":2C50
      Left            =   5280
      List            =   "alta_proveedor.frx":2C96
      TabIndex        =   4
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox Text4 
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
      Left            =   1560
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text3 
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
      Left            =   1680
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text2 
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2280
      Width           =   7335
   End
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
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label9 
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
      Left            =   7920
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ID :"
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
      Left            =   7320
      TabIndex        =   17
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Correo Electronico:"
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
      Left            =   3240
      TabIndex        =   15
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia:"
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
      Left            =   3840
      TabIndex        =   13
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre / Razon Social:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Altas de Proveedores"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "alta_proveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
    MsgBox "Todos los campos son obligatorios", vbOKOnly, " VIVERO SAN NICOLAS"
Else
    g = "insert into proveedores values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "','" & Text4 & "','" & Text5 & "'," & Val(Label9) & ")"
    conexion_BD.Execute g
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Label9 = Val(Label9) + 1
End If

MSFlexGrid1.Clear
MSFlexGrid1.Visible = False
interno = "select max(id_prove) from proveedores"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Label9 = 2500
Else
    Label9 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub

Private Sub Command2_Click()
MSFlexGrid1.Visible = True
Command3.Visible = True
Command4.Visible = True
Command1.Visible = False
LISTA = "select * from proveedores order by nombre_prove"
TABLA.Open LISTA, conexion_BD
lin = 0
With MSFlexGrid1
Do While Not TABLA.EOF = True
    lin = lin + 1
    .Rows = MSFlexGrid1.Rows + 1
    .TextMatrix(lin, 0) = TABLA!nombre_prove
    .TextMatrix(lin, 1) = TABLA!direccion
    .TextMatrix(lin, 2) = TABLA!ciudad
    .TextMatrix(lin, 3) = TABLA!provincia
    .TextMatrix(lin, 4) = TABLA!telefono
    .TextMatrix(lin, 5) = TABLA!email
    .TextMatrix(lin, 6) = TABLA!id_prove
    TABLA.MoveNext
    Loop
TABLA.Close
End With
End Sub

Private Sub Command3_Click()
resp = MsgBox("Está a punto de modificar datos importantes!!! Desea continuar?", vbYesNo)
If resp = vbYes Then

    h = "update proveedores set nombre_prove='" & Text1 & "',direccion='" & Text2 & "',ciudad='" & Text3 & "',provincia='" & Combo1 & "',telefono='" & Text4 & "',email='" & Text5 & "' where id_prove= " & Val(Label9) & ""
    conexion_BD.Execute h
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
Else
    Text1.SetFocus
End If
interno = "select max(id_prove) from proveedores"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Label9 = 2500
Else
    Label9 = TABLA.Fields(0) + 1
End If
TABLA.Close
Command1.Visible = True
Command3.Visible = False

End Sub

Private Sub Command4_Click()

ques = MsgBox("Al eliminar un proveedor se elimina de forma permanente. Desea continuar?", vbYesNo)
If ques = vbYes Then
    Borrar = "delete * from proveedores where nombre_prove='" & Text1 & "'"
    conexion_BD.Execute Borrar
    MsgBox "El cliente ha sido borrado."
    Command4.Visible = False
    
End If

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
MSFlexGrid1.Clear
MSFlexGrid1.Visible = False
interno = "select max(id_prove) from proveedores"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Label9 = 2500
Else
    Label9 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub

Private Sub Command5_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
MSFlexGrid1.Clear
MSFlexGrid1.Visible = False
interno = "select max(id_prove) from proveedores"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Label9 = 2500
Else
    Label9 = TABLA.Fields(0) + 1
End If
TABLA.Close
Command4.Visible = False
Command1.Visible = True
End Sub



Private Sub Form_Load()

interno = "select max(id_prove) from proveedores"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Label9 = 2500
Else
    Label9 = TABLA.Fields(0) + 1
End If
TABLA.Close

With MSFlexGrid1
.FixedCols = 0
.Cols = 7
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "NOMBRE / RAZON SOCIAL"
.TextMatrix(0, 1) = "DIRECCION"
.TextMatrix(0, 2) = "LOCALIDAD"
.TextMatrix(0, 3) = "PROVINCIA"
.TextMatrix(0, 4) = "TELEFONO"
.TextMatrix(0, 5) = "CORREO ELECTRONICO"
.TextMatrix(0, 6) = "ID"


.ColWidth(0) = 3000
.ColWidth(1) = 1000
.ColWidth(2) = 3000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 500
End With
End Sub

Private Sub MSFlexGrid1_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text1 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0) ' nombre
Text2 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1) 'direccion
Text3 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2) 'localidad
Combo1 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3) 'provincia
Text4 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4) 'telefono
Text5 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 5) 'email
Label9 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 6) 'id

End Sub

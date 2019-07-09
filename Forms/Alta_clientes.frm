VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Alta_clientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Altas de Clientes"
   ClientHeight    =   10290
   ClientLeft      =   2745
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "Alta_clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   9600
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   240
      TabIndex        =   32
      Top             =   4200
      Visible         =   0   'False
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
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
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   31
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000A&
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000A&
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   29
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      DisabledPicture =   "Alta_clientes.frx":058A
      DragIcon        =   "Alta_clientes.frx":0B14
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
      Left            =   2880
      Picture         =   "Alta_clientes.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Index           =   0
      Left            =   9000
      Picture         =   "Alta_clientes.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   11640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      DisabledPicture =   "Alta_clientes.frx":1BB2
      DragIcon        =   "Alta_clientes.frx":213C
      Height          =   495
      Index           =   0
      Left            =   -4080
      Picture         =   "Alta_clientes.frx":26C6
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar Cambios"
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
      Index           =   1
      Left            =   4200
      Picture         =   "Alta_clientes.frx":2C50
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5760
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
      Left            =   240
      Picture         =   "Alta_clientes.frx":31DA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command2 
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
      Index           =   1
      Left            =   1560
      Picture         =   "Alta_clientes.frx":3764
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2880
      TabIndex        =   11
      Top             =   5040
      Width           =   6255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5520
      TabIndex        =   10
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "Alta_clientes.frx":3CEE
      Left            =   1680
      List            =   "Alta_clientes.frx":3D34
      TabIndex        =   9
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7080
      TabIndex        =   8
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   7
      Top             =   3600
      Width           =   3975
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "Alta_clientes.frx":3E19
      Left            =   7680
      List            =   "Alta_clientes.frx":3E23
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "Alta_clientes.frx":3E2F
      Left            =   4320
      List            =   "Alta_clientes.frx":3E3C
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   5655
   End
   Begin VB.CommandButton Command5 
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
      Left            =   240
      Picture         =   "Alta_clientes.frx":3E66
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el Nº de la empresa:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Factura a nombre de empresa?"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6960
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Productor:"
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
      TabIndex        =   27
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Correo Electrónico:"
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
      TabIndex        =   23
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label9 
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
      Left            =   4200
      TabIndex        =   22
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label8 
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
      Left            =   360
      TabIndex        =   21
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label7 
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
      Left            =   5760
      TabIndex        =   20
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      Left            =   360
      TabIndex        =   19
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cta Cte:"
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
      Left            =   6600
      TabIndex        =   18
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "I.V.A.:"
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
      Left            =   3480
      TabIndex        =   17
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "C.U.I.T.:"
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
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre / Razón Social: "
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
      TabIndex        =   15
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Alta_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    MsgBox "Todos los campos son obligatorios!", vbOKOnly, "VIVERO SAN NICOLAS SA"
Else
    a = "insert into clientes values ('" & Text1 & "','" & Text2 & "','" & Combo1 & "','" & Combo2 & "','" & Text3 & "','" & Text4 & "','" & Combo3 & "','" & Text5 & "','" & Text6 & "'," & Val(Text7) & "," & Val(Text8) & ")"
    conexion_BD.Execute a
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = Val(Text7) + 1
    Text8 = ""
    
End If

Alta_clientes.Refresh
End Sub

Private Sub Command2_Click(Index As Integer)
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Command1.Visible = True
Command3(1).Visible = False
interno = "select max(id_cliente) from clientes"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text7 = 1000
Else
    Text7 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub


Private Sub Command3_Click(Index As Integer)
resp = MsgBox("Esta por modificar datos importantes!!! Esta seguro de continuar?", vbYesNo)
If resp = vbYes Then

    modif = " update clientes set CUIT= '" & Text2 & "', IVA='" _
    & Combo1 & "', CTACTE='" & Combo2 & "', direccion='" & Text3 & "',localidad='" _
    & Text4 & "',provincia='" & Combo3 & "',telefono='" & Text5 & "',email='" _
    & Text6 & "' where id_cliente=" & Val(Text7) & ""
    conexion_BD.Execute modif
Else
    Text2.SetFocus
End If
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Command1.Visible = True
Command3(1).Visible = False
Command5.Visible = False
interno = "select max(id_cliente) from clientes"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text7 = 1000
Else
    Text7 = TABLA.Fields(0) + 1
End If
TABLA.Close
Command3(1).Visible = False
Command1.Visible = True
End Sub

Private Sub Command4_Click()
Command3(1).Visible = True
Command5.Visible = True
Command1.Visible = False
c = "select * from clientes order by nombre_cliente"
TABLA.Open c, conexion_BD
Call ALTA_GRID
Frame1.Visible = True
Frame1.BorderStyle = 0
lin = 0
With MSFlexGrid1
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!nombre_cliente
    .TextMatrix(lin, 1) = TABLA!CUIT
    .TextMatrix(lin, 2) = TABLA!iva
    .TextMatrix(lin, 3) = TABLA!CTACTE
    .TextMatrix(lin, 4) = TABLA!direccion
    .TextMatrix(lin, 5) = TABLA!localidad
    .TextMatrix(lin, 6) = TABLA!provincia
    .TextMatrix(lin, 7) = TABLA!telefono
    .TextMatrix(lin, 8) = TABLA!email
    .TextMatrix(lin, 9) = TABLA!id_cliente
    TABLA.MoveNext
    Loop
TABLA.Close
End With
End Sub

Private Sub Command5_Click()
ques = MsgBox("Al eliminar un Cliente se elimina de forma permanente. Desea continuar?", vbYesNo)
If ques = vbYes Then
    Borrar = "delete * from clientes where nombre_cliente='" & Text1 & "'"
    conexion_BD.Execute Borrar
    MsgBox "El cliente ha sido borrado."
    Command5.Visible = False
    
End If

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Command1.Visible = True
Command3(1).Visible = False
interno = "select max(id_cliente) from clientes"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text7 = 1000
Else
    Text7 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub

Private Sub Form_Activate()
interno = "select max(id_cliente) from clientes"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text7 = 1000
Else
    Text7 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub

Private Sub Form_Load()
interno = "select max(id_cliente) from clientes"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text7 = 1000
Else
    Text7 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub



Private Sub MSFlexGrid1_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text1 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0) ' nombre
Text2 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1) ' cuit
Combo1 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2) 'IVA
Combo2 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3) 'CTA CTE
Text3 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4) 'direccion
Text4 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 5) 'localidad
Combo3 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 6) 'provincia
Text5 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 7) 'telefono
Text6 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 8) 'email
Text7 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 9) 'cuit
Frame1.Visible = False
End Sub

Private Sub Option1_Click()
If Option1 = True Then
    Label12.Visible = True
    Text8.Visible = True
Else
    Label12.Visible = False
    Text8.Visible = False
    Text8 = "0"
End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
    Label12.Visible = False
    Text8.Visible = False
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    
Else
    Label12.Visible = True
    Text8.Visible = True
End If
End Sub



Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    c = "select * from clientes_empresa where id_empresa= " & Val(Text8) & ""
TABLA.Open c, conexion_BD

If TABLA.EOF = True Then
    res = MsgBox("Nº no encontrado. Intente nuevamente", vbOKOnly, "VIVERO SAN NICOLAS")
Else
Text2 = TABLA!CUIT
Combo1 = TABLA!iva
Combo2 = TABLA!CTACTE
Text3 = TABLA!direccion
Text4 = TABLA!localidad
Combo3 = TABLA!provincia
Text5 = TABLA!telefono
Text6 = TABLA!email

TABLA.Close
Command3(1).Visible = False
Command1.Visible = True
End If
End If
End Sub
Private Sub ALTA_GRID()
With MSFlexGrid1
.FixedCols = 0
.Cols = 10
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "NOMBRE / RAZON SOCIAL"
.TextMatrix(0, 1) = "CUIT"
.TextMatrix(0, 2) = "IVA"
.TextMatrix(0, 3) = "CTACTE"
.TextMatrix(0, 4) = "DIRECCION"
.TextMatrix(0, 5) = "LOCALIDAD"
.TextMatrix(0, 6) = "PROVINCIA"
.TextMatrix(0, 7) = "TELEFONO"
.TextMatrix(0, 8) = "CORREO ELECTRONICO"
.TextMatrix(0, 9) = "ID"


.ColWidth(0) = 3000
.ColWidth(1) = 1000
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 3000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 500
End With
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Alta_empresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alta de Empresas"
   ClientHeight    =   10290
   ClientLeft      =   2820
   ClientTop       =   435
   ClientWidth     =   9630
   Icon            =   "Alta_empresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   9630
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   240
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   25
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      DisabledPicture =   "Alta_empresa.frx":058A
      DragIcon        =   "Alta_empresa.frx":0B14
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      Picture         =   "Alta_empresa.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text7 
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
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   23
      Top             =   2400
      Width           =   855
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
      Left            =   3480
      TabIndex        =   0
      Top             =   1680
      Width           =   5895
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
      Left            =   4800
      MaxLength       =   11
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
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
      ItemData        =   "Alta_empresa.frx":1628
      Left            =   7440
      List            =   "Alta_empresa.frx":1635
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Alta_empresa.frx":165F
      Left            =   1440
      List            =   "Alta_empresa.frx":1669
      TabIndex        =   3
      Top             =   3120
      Width           =   975
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
      Left            =   3840
      TabIndex        =   4
      Top             =   3120
      Width           =   5535
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
      Left            =   6480
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "Alta_empresa.frx":1675
      Left            =   6360
      List            =   "Alta_empresa.frx":16BB
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
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
      Left            =   1680
      TabIndex        =   7
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text6 
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
      TabIndex        =   8
      Top             =   5280
      Width           =   6495
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
      Height          =   615
      Index           =   1
      Left            =   1680
      Picture         =   "Alta_empresa.frx":17A0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   855
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
      Height          =   615
      Left            =   360
      Picture         =   "Alta_empresa.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   855
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
      Height          =   615
      Index           =   1
      Left            =   4320
      Picture         =   "Alta_empresa.frx":22B4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresas"
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
      TabIndex        =   22
      Top             =   360
      Width           =   2295
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
      TabIndex        =   21
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "C.U.I.T.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "I.V.A.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6600
      TabIndex        =   19
      Top             =   2400
      Width           =   855
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
      Left            =   360
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
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
      Left            =   5160
      TabIndex        =   16
      Top             =   3960
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
      Left            =   9840
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
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
      Left            =   360
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
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
      TabIndex        =   13
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Empresa:"
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
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "Alta_empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    MsgBox "Todos los campos son obligatorios!", vbOKOnly, "VIVERO SAN NICOLAS SA"
Else
    a = "insert into clientes_empresa values ('" & Text1 & "','" & Text2 & "','" & Combo1 & "','" & Combo2 & "','" & Text3 & "','" & Text4 & "','" & Combo3 & "','" & Text5 & "','" & Text6 & "'," & Val(Text7) & ")"
    conexion_BD.Execute a
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    'Text7 = ""
    Text7 = Val(Text7) + 1
    
End If
Alta_empresa.Refresh ' no me vuelve a cargar los datos de nuevo
End Sub

Private Sub Command2_Click(Index As Integer)
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Command1.Visible = True
Command3(1).Visible = False
inter = "select max(id_empresa) from clientes_empresa"
TABLA.Open inter, conexion_BD
If TABLA.EOF = True Then
    Text7 = 100
Else
    Text7 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub

Private Sub Command3_Click(Index As Integer)
resp = MsgBox("Esta por modificar datos importantes!!! Esta seguro de continuar?", vbYesNo)
If resp = vbYes Then

    modif = " update clientes_empresa set CUIT= '" & Text2 & "', IVA='" _
    & Combo1 & "', CTACTE='" & Combo2 & "', direccion='" & Text3 & "',localidad='" _
    & Text4 & "',provincia='" & Combo3 & "',telefono='" & Text5 & "',email='" _
    & Text6 & "' where id_empresa=" & Val(Text7) & ""
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
interno = "select max(id_empresa) from clientes_empresa"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text7 = 1000
Else
    Text7 = TABLA.Fields(0) + 1
End If
TABLA.Close

End Sub

Private Sub Command4_Click()
c = "select * from clientes_empresa where id_empresa= " & Val(Text7) & ""
TABLA.Open c, conexion_BD

If TABLA.EOF = True Then
    res = MsgBox("Nº no encontrado. Desea cargarlo?", vbYesNo, "VIVERO SAN NICOLAS SA")
    If res = vbNo Then
        TABLA.Close
        Text1 = ""
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Text6 = ""
        Text7 = ""
        Text1.SetFocus
        Command1.Visible = True
        Command3(1).Visible = False
    Else
        Text2.SetFocus
    End If
Else
Text1 = TABLA!nombre_empresa
Text2 = TABLA!CUIT
Combo1 = TABLA!iva
Combo2 = TABLA!CTACTE
Text3 = TABLA!direccion
Text4 = TABLA!localidad
Combo3 = TABLA!provincia
Text5 = TABLA!telefono
Text6 = TABLA!email
Text7 = TABLA!id_empresa
TABLA.Close
Command3(1).Visible = True
Command1.Visible = False

End If

End Sub

Private Sub Command5_Click()
Command3(1).Visible = True
Command1.Visible = False
c = "select * from clientes_empresa order by nombre_empresa"
TABLA.Open c, conexion_BD
Call ALTA_GRID
Frame1.Visible = True
Frame1.BorderStyle = 0
lin = 0
With MSFlexGrid1
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = MSFlexGrid1.Rows + 1
    .TextMatrix(lin, 0) = TABLA!nombre_empresa
    .TextMatrix(lin, 1) = TABLA!CUIT
    .TextMatrix(lin, 2) = TABLA!iva
    .TextMatrix(lin, 3) = TABLA!CTACTE
    .TextMatrix(lin, 4) = TABLA!direccion
    .TextMatrix(lin, 5) = TABLA!localidad
    .TextMatrix(lin, 6) = TABLA!provincia
    .TextMatrix(lin, 7) = TABLA!telefono
    .TextMatrix(lin, 8) = TABLA!email
    .TextMatrix(lin, 9) = TABLA!id_empresa
    TABLA.MoveNext
    Loop
TABLA.Close
End With
End Sub

Private Sub Form_Activate()
inter = "select max(id_empresa) from clientes_empresa"
TABLA.Open inter, conexion_BD
If TABLA.EOF = True Then
    Text7 = 2000
Else
    Text7 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub

Private Sub Form_Load()
inter = "select max(id_empresa) from clientes_empresa"
TABLA.Open inter, conexion_BD
If TABLA.EOF = True Then
    Text7 = 2000
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

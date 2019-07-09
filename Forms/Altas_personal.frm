VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Altas_personal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Altas de Personal"
   ClientHeight    =   9150
   ClientLeft      =   2820
   ClientTop       =   495
   ClientWidth     =   7605
   Icon            =   "Altas_personal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   7605
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   360
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   7095
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5055
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8916
         _Version        =   393216
         BackColor       =   16777152
         SelectionMode   =   1
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
      Left            =   6000
      Picture         =   "Altas_personal.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      DisabledPicture =   "Altas_personal.frx":0B14
      DragIcon        =   "Altas_personal.frx":109E
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
      Left            =   4800
      Picture         =   "Altas_personal.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
      Width           =   975
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
      Left            =   5280
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
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
      Left            =   4800
      Picture         =   "Altas_personal.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
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
      Left            =   6000
      Picture         =   "Altas_personal.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   9
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
      Left            =   1800
      TabIndex        =   7
      Top             =   5760
      Width           =   2175
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
      ItemData        =   "Altas_personal.frx":26C6
      Left            =   1800
      List            =   "Altas_personal.frx":270C
      TabIndex        =   6
      Top             =   5040
      Width           =   2175
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
      Left            =   1800
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
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
      Left            =   1800
      TabIndex        =   4
      Top             =   3600
      Width           =   5295
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
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
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
      Left            =   2880
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
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
      Left            =   4800
      Picture         =   "Altas_personal.frx":27F1
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "C.U.I.L:"
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
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label7 
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
      TabIndex        =   15
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label5 
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
      Left            =   360
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Legajo Nº:"
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
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido y Nombre:"
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
      TabIndex        =   10
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Altas del Personal"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "Altas_personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    MsgBox "Todos los campos son obligatorios!", vbOKOnly, "VIVERO SAN NICOLAS SA"
Else
    'validar = " select * from alta_rrhh where id_rrhh=" & Val(Text6) & ""
    'TABLA.Open validar, conexion_BD
    'If existe = True Then
    '    MsgBox "El Nº de legajo seleccionado ya ha sido registrado"
    'Else
        a = "insert into alta_rrhh values (" & Val(Text6) & ",'" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Combo1 & "','" & Text5 & "')"
        conexion_BD.Execute a
    'End If
End If
    Text6 = ""
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""

End Sub

Private Sub Command2_Click()
Text6 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
interno = "select max(id_rrhh) from alta_rrhh"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text6 = 1
Else
    Text6 = TABLA.Fields(0) + 1
End If
TABLA.Close
Command1.Visible = True
Command5.Visible = False
End Sub

Private Sub Command3_Click(Index As Integer)
resp = MsgBox("Esta a punto de modificar datos importantes!!! Desea continuar?", vbYesNo)
If resp = vbYes Then

    h = "update alta_rrhh set id_rrhh=" & Val(Text6) & ",cuil='" & Text2 & "',direccion='" & Text3 & "',provincia='" & Combo1 & "',ciudad='" & Text4 & "',telefono='" & Text5 & "' where nombre_rrhh= '" & Text1 & "'"
    conexion_BD.Execute h
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
Else
    Text1.SetFocus
End If
interno = "select max(id_rrhh) from alta_rrhh"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text6 = 1
Else
    Text6 = TABLA.Fields(0) + 1
End If
TABLA.Close
Command3(1).Visible = False
Command1.Visible = True
End Sub

Private Sub Command4_Click()
Frame1.Visible = True
W = "select * from alta_rrhh order by nombre_rrhh asc"
TABLA.Open W, conexion_BD

Call alta_rrhh
With MSFlexGrid1
    Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!id_rrhh
        .TextMatrix(lin, 1) = TABLA!nombre_rrhh
        .TextMatrix(lin, 2) = TABLA!cuil
        .TextMatrix(lin, 3) = TABLA!direccion
        .TextMatrix(lin, 4) = TABLA!ciudad
        .TextMatrix(lin, 5) = TABLA!provincia
        .TextMatrix(lin, 6) = TABLA!telefono
        TABLA.MoveNext
    Loop
    
    TABLA.Close
End With
Command3(1).Visible = True
Command1.Visible = False
Command5.Visible = True
End Sub
Private Sub alta_rrhh()

With MSFlexGrid1
.FixedCols = 0
.Cols = 7
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "LEGAJO Nº"
.TextMatrix(0, 1) = "APELLIDO Y NOMBRE"
.TextMatrix(0, 2) = "CUIL"
.TextMatrix(0, 3) = "DIRECCION"
.TextMatrix(0, 4) = "LOCALIDAD"
.TextMatrix(0, 5) = "PROVINCIA"
.TextMatrix(0, 6) = "TELEFONO"

.ColWidth(0) = 1000
.ColWidth(1) = 3000
.ColWidth(2) = 2000
.ColWidth(3) = 3000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
End With
End Sub

Private Sub Command5_Click()
ques = MsgBox("Al eliminar a un personal se elimina de forma permanente. Desea continuar?", vbYesNo)
If ques = vbYes Then
    Borrar = "delete * from alta_rrhh where nombre_rrhh='" & Text1 & "'"
    conexion_BD.Execute Borrar
    MsgBox "El personal ha sido borrado."
    Command5.Visible = False
    
End If

Text6 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
interno = "select max(id_rrhh) from alta_rrhh"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text6 = 1
Else
    Text6 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub



Private Sub MSFlexGrid1_Click()
datos = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0) ' toma el legajo!

buscar = "select * from alta_rrhh where id_rrhh= " & Val(datos) & ""
TABLA.Open buscar, conexion_BD
nombre = TABLA!nombre_rrhh
cuil = TABLA!cuil
direccion = TABLA!direccion
localidad = TABLA!ciudad
prov = TABLA!provincia
telefono = TABLA!telefono
TABLA.Close
Frame1.Visible = False

Text1 = nombre
Text2 = cuil
Text3 = direccion
Text4 = localidad
Combo1 = prov
Text5 = telefono
Text6 = datos
End Sub

Private Sub Form_Load()
interno = "select max(id_rrhh) from alta_rrhh"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Text6 = 1
Else
    Text6 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub

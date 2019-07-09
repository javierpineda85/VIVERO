VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Alta_usuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Altas y modificación de usuarios"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   Icon            =   "Alta_usuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   360
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Volver"
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
         Left            =   2040
         Picture         =   "Alta_usuarios.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2640
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2295
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4048
         _Version        =   393216
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
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Info"
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
      Left            =   4200
      Picture         =   "Alta_usuarios.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar"
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
      Left            =   5520
      MaskColor       =   &H00808080&
      Picture         =   "Alta_usuarios.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   975
   End
   Begin VB.ComboBox nivel_usua 
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
      ItemData        =   "Alta_usuarios.frx":1628
      Left            =   2880
      List            =   "Alta_usuarios.frx":163B
      TabIndex        =   12
      Top             =   4080
      Width           =   975
   End
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
      Left            =   5520
      Picture         =   "Alta_usuarios.frx":164E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listar"
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
      Left            =   5520
      Picture         =   "Alta_usuarios.frx":1BD8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
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
      Left            =   5520
      Picture         =   "Alta_usuarios.frx":2162
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
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
      Left            =   5520
      Picture         =   "Alta_usuarios.frx":26EC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox contrasena 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox usuario 
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
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel de Usuario:"
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
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Repetir contraseña:"
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
      TabIndex        =   3
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
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
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo usuario:"
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
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión de Usuarios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "Alta_usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If contrasena.Text = Text3.Text Then
    Guardar = "insert into usuarios values('" & usuario & "','" & contrasena & "','" & nivel_usua & "')"
    conexion_BD.Execute Guardar
    MsgBox "El usuario ha sido registrado exitosamente"
Else
    MsgBox "Las contraseñas ingresadas no coinciden"
    contrasena = ""
    Text3 = ""
End If
usuario = ""
contrasena = ""
Text3 = ""
End Sub

Private Sub Command2_Click()
If contrasena.Text = Text3.Text Then
    modif = "update usuarios set contraseña='" & contrasena & "',nivel_usua='" & nivel_usua & "' where usuario='" & usuario & "'"
    conexion_BD.Execute modif
    MsgBox "El usuario ha sido modificado exitosamente"
Else
    MsgBox "Las contraseñas ingresadas no coinciden"
    contrasena = ""
    Text3 = ""
End If
usuario = ""
contrasena = ""
Text3 = ""

End Sub

Private Sub Command3_Click()
Frame1.Visible = True
Command2.Visible = True
Command1.Visible = False
Call LISTA

buscar = "select * from usuarios order by usuario"
TABLA.Open buscar, conexion_BD

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!usuario
    If TABLA!usuario = "ADMIN" Then
        MSFlexGrid1.Row = lin
        MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
        lin = lin - 1
    Else
    
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!nivel_usua
    MSFlexGrid1.TextMatrix(lin, 2) = TABLA!contraseña
    End If
    TABLA.MoveNext
Loop
TABLA.Close
End Sub

Private Sub Command4_Click()
ques = MsgBox("Al eliminar un usuario se elimina permanentemente. Desea continuar?", vbYesNo)
If ques = vbYes Then
    Borrar = "delete * from usuarios where usuario='" & usuario & "'"
    conexion_BD.Execute Borrar
    MsgBox "El usuario ha sido borrado"

End If
usuario = ""
contrasena = ""
Text3 = ""
End Sub
Private Sub LISTA()
With MSFlexGrid1
.FixedCols = 0
.Cols = 3
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "USUARIO"
.TextMatrix(0, 1) = "NIVEL"
.TextMatrix(0, 2) = "contra"

.ColWidth(0) = 2500
.ColWidth(1) = 800
.ColWidth(2) = 0
End With
End Sub

Private Sub Command5_Click()
Frame1.Visible = False
Command2.Visible = False
Command1.Visible = True
End Sub

Private Sub Command6_Click()
usuario = ""
contrasena = ""
Text3 = ""
End Sub

Private Sub Command7_Click()
'ver informacion de usuarios
Info_nivel_usuario.Show
End Sub

Private Sub MSFlexGrid1_Click()
usuario = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
contrasena = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
Text3 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
nivel_usua = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
Frame1.Visible = False
End Sub

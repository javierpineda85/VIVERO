VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Reedicion_facturas 
   Caption         =   "Gestión de Modificación de Facturas"
   ClientHeight    =   8040
   ClientLeft      =   1830
   ClientTop       =   555
   ClientWidth     =   10290
   Icon            =   "Reedicion_facturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10290
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   1080
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   6975
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
         Left            =   5400
         Picture         =   "Reedicion_facturas.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command4 
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
         Left            =   3240
         Picture         =   "Reedicion_facturas.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4320
         Picture         =   "Reedicion_facturas.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2040
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61800449
         CurrentDate     =   41177
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
         Left            =   1200
         TabIndex        =   20
         Top             =   1560
         Width           =   5415
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
         Left            =   4800
         TabIndex        =   19
         Top             =   960
         Width           =   1815
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
         Left            =   1800
         TabIndex        =   18
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Left            =   2760
         TabIndex        =   26
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha:"
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
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre/Razón social:"
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
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Factura:"
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
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
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
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
   End
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
      Left            =   8880
      Picture         =   "Reedicion_facturas.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
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
      Left            =   8880
      Picture         =   "Reedicion_facturas.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   2535
      Begin VB.OptionButton OpCli 
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
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton OpPer 
         Caption         =   "Personal"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton OpPro 
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
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4215
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command7 
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
         Left            =   3600
         Picture         =   "Reedicion_facturas.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3360
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5106
         _Version        =   393216
         BackColor       =   16777152
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   4575
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label21 
         Caption         =   "Seleccione el Nombre/Razón Social que desea buscar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Modificación de Facturas"
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
      Left            =   2520
      TabIndex        =   17
      Top             =   360
      Width           =   5295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Reedicion_facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As String
Private Sub FACTURAS_GRID()
With MSFlexGrid2

.FixedCols = 0
.Cols = 4
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "FACTURA"
.TextMatrix(0, 1) = "FECHA"
.TextMatrix(0, 2) = "MONTO"
.TextMatrix(0, 3) = "DETALLE"

.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 4000
End With
End Sub


Private Sub Combo2_Click()
If op = "PROVEEDORES" Then
    SQL = "select * from prove_a_pagar where proveedor='" & Combo2 & "'"
    TABLA.Open SQL, conexion_BD
    Call FACTURAS_GRID
    With MSFlexGrid2
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!factura
        .TextMatrix(lin, 1) = TABLA!fecha
        If IsNull(TABLA!monto) Then
            .TextMatrix(lin, 2) = "SIN DATOS"
        Else
            .TextMatrix(lin, 2) = TABLA!monto
        End If
        If IsNull(TABLA!detalle) Then
            .TextMatrix(lin, 3) = "SIN DATOS"
        Else
            .TextMatrix(lin, 3) = TABLA!detalle
        End If
        TABLA.MoveNext
    Loop
    TABLA.Close
    End With
Else 'op="CLIENTES'
    SQL = "select * from clientes_a_pagar where cliente='" & Combo2 & "'"
    TABLA.Open SQL, conexion_BD
    Call FACTURAS_GRID
    With MSFlexGrid2
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!factura
        .TextMatrix(lin, 1) = TABLA!fecha
        If IsNull(TABLA!monto) Then
            .TextMatrix(lin, 2) = "SIN DATOS"
        Else
            .TextMatrix(lin, 2) = TABLA!monto
        End If
        If IsNull(TABLA!detalle) Then
            .TextMatrix(lin, 3) = "SIN DATOS"
        Else
            .TextMatrix(lin, 3) = TABLA!detalle
        End If
        TABLA.MoveNext
    Loop
    TABLA.Close
    End With
End If

Frame4.Visible = True

End Sub

Private Sub Command2_Click()
Label2 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Frame3.Visible = False
Frame4.Visible = False


End Sub

Private Sub Command3_Click()
Frame3.Visible = False
End Sub

Private Sub Command4_Click()
If op = "CLIENTES" Then
    modif = " update clientes_a_pagar set fecha='" & DTPicker1 & "', monto='" _
    & Text3 & "', detalle='" & Text4 & "'where factura='" & Text2 & "' and cliente='" & Label2 & "'"
    conexion_BD.Execute modif
Else
    modif = "update prove_a_pagar set fecha='" & DTPicker1 & "', monto='" _
    & Text3 & "', detalle='" & Text4 & "' where factura='" & Text2 & "' and proveedor='" & Label2 & "'"
    conexion_BD.Execute modif
End If
Frame3.Visible = False
Frame4.Visible = False
Label2 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command5_Click()
If op = "CLIENTES" Then
    ques = MsgBox("Esta por eliminar de manera permanente una factura, desea continuar?", vbYesNo)
    If ques = vbYes Then
        SQL = " delete * from clientes_a_pagar where factura='" & Text2 & "' and cliente='" & Label4 & "'"
        conexion_BD.Execute SQL
    End If
Else
    ques = MsgBox("Esta por eliminar de manera permanente una factura, desea continuar?", vbYesNo)
    If ques = vbYes Then
        SQL = " delete * from prove_a_pagar where factura='" & Text2 & "' and proveedor='" & Label4 & "'"
        conexion_BD.Execute SQL
    End If
End If
Frame3.Visible = False
Frame4.Visible = False
Label2 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command7_Click()
Frame4.Visible = False
End Sub



Private Sub MSFlexGrid2_Click()
Label2 = Combo2.Text
Text2 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 0)
DTPicker1 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 1)
Text3 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 2)
Text4 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 3)

Frame3.Visible = True

End Sub

Private Sub OpCli_Click()
If OpCli = True Then
    Combo2.Clear
    op = "CLIENTES"
    SQL = "select * from clientes order by nombre_cliente"
    TABLA.Open SQL, conexion_BD
    
    Do While Not TABLA.EOF
        Combo2.AddItem TABLA!nombre_cliente
        
        TABLA.MoveNext
    Loop
    TABLA.Close
End If
Frame2.Visible = True
Frame2.BorderStyle = 0
End Sub


Private Sub OpPro_Click()
If OpPro = True Then
    op = "PROVEEDORES"
    Combo2.Clear
    SQL = " select * from proveedores order by nombre_prove"
    TABLA.Open SQL, conexion_BD
    
    Do While Not TABLA.EOF
        Combo2.AddItem TABLA!nombre_prove
        TABLA.MoveNext
    Loop
    TABLA.Close
End If
Frame2.Visible = True
Frame2.BorderStyle = 0

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Reimpresiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de reimpresion"
   ClientHeight    =   9480
   ClientLeft      =   1710
   ClientTop       =   540
   ClientWidth     =   10140
   Icon            =   "Reimpresiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   10140
   Begin VB.Frame Frame5 
      Caption         =   "Personal"
      Height          =   1575
      Left            =   5400
      TabIndex        =   35
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
         TabIndex        =   37
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label21 
         Caption         =   "Seleccione el nombre del personal que desea buscar:"
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
         TabIndex        =   36
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   5895
      Left            =   240
      TabIndex        =   31
      Top             =   3240
      Visible         =   0   'False
      Width           =   8655
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
         Height          =   615
         Left            =   7560
         Picture         =   "Reimpresiones.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1680
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
         Left            =   5880
         TabIndex        =   38
         Top             =   1080
         Width           =   2535
      End
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
         Height          =   615
         Left            =   2520
         Picture         =   "Reimpresiones.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5160
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4815
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   8493
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
      Begin VB.Label Label22 
         Caption         =   "Buscar por Nombre:"
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
         Left            =   5880
         TabIndex        =   39
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
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
      Left            =   2880
      TabIndex        =   27
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
      Begin VB.OptionButton Option5 
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
         TabIndex        =   30
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option4 
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
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
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
         TabIndex        =   28
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   360
      TabIndex        =   24
      Top             =   1560
      Width           =   2535
      Begin VB.OptionButton Option2 
         Caption         =   "Nº de Remito interno"
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
         TabIndex        =   26
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cuentas Corrientes"
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
         TabIndex        =   25
         Top             =   480
         Width           =   2055
      End
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
      Picture         =   "Reimpresiones.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4320
      Width           =   975
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
      Picture         =   "Reimpresiones.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   6360
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   8775
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1815
         Left            =   240
         TabIndex        =   21
         Top             =   2880
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   16777152
         BackColorBkg    =   -2147483633
         GridLinesFixed  =   3
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
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   20
         Top             =   5400
         Width           =   3735
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
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
         TabIndex        =   19
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto total en $"
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
         TabIndex        =   18
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
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
         Left            =   7320
         TabIndex        =   17
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Retenciones:"
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
         TabIndex        =   16
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
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
         Left            =   4320
         TabIndex        =   15
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Efectivo: $"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Según el siguiente detalle:"
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
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
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
         TabIndex        =   12
         Top             =   1800
         Width           =   6255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "En concepto de:"
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
         TabIndex        =   11
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
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
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "la cantidad de pesos $"
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
         TabIndex        =   9
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
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
         TabIndex        =   8
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Recibimos de:"
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
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "fecha"
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
         Left            =   6840
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de recibo:"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "L13 "
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
      Left            =   8160
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione un remito para imprimir:"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión de Reimpresión"
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
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Reimpresiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CTACTE As String
Private Sub Combo1_Click()
Call ALTAGRID_CHEQUES
buscar = "select * from remito_interno where n_remito =" & Combo1.Text & ""
TABLA.Open buscar, conexion_BD

Label4 = TABLA!n_remito
Label5 = TABLA!fecha
Label7 = TABLA!origen
Label9 = TABLA!total
Label11 = TABLA!detalle
Label15 = "0"
Label17 = "0"
Label19 = TABLA!total
Label20 = TABLA!destino
TABLA.Close
Frame1.Visible = True

Select Case Label7 ' origen

Case "VIVERO SAN NICOLAS S.A."

        ''' DEL MOV_PROVEEDOR TRAE EL PAGO EN EFECTIVO'''
        prove = "select * from mov_proveedor where r_interno='" & Label4 & "'"
        TABLA.Open prove, conexion_BD
        Do While Not TABLA.EOF
            If TABLA!n_cheque = "0" Then
                Label15 = TABLA!pago
            End If
            TABLA.MoveNext
        Loop
        TABLA.Close
        ''' DEL SALECHEQUE TRAE LOS DATOS DE LOS CHEQUES QUE SE HAYAN USADO'''

        listche = "select * from mov_cheques where fecha_mod ='" & Label5 & "' and destino ='" & Label20 & "'"
        TABLA.Open listche, conexion_BD
        
        Call ALTAGRID_CHEQUES
        Do While Not TABLA.EOF
            lin = lin + 1
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
            MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
            MSFlexGrid1.TextMatrix(lin, 2) = TABLA!importe
            'MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
            MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_vto
            If TABLA!importe < "0" Then
                importe = CDbl(MSFlexGrid1.TextMatrix(lin, 2)) * -1
                MSFlexGrid1.TextMatrix(lin, 2) = importe
            End If

            TABLA.MoveNext
        Loop
        TABLA.Close
        
        ''' COMO RETENCIONES NO HAY COMPLETAMOS EL REMITO''
        Label17 = ""

    
''' DEL MOV_CLIENTES TRAE EL PAGO EN EFECTIVO'''
cliente = "select * from mov_clientes where r_interno='" & Label4 & "'"
TABLA.Open cliente, conexion_BD
Do While Not TABLA.EOF
    If TABLA!n_cheque = "0" Then
        Label15 = TABLA!pago
    End If
    TABLA.MoveNext
Loop
TABLA.Close

Case Else

''' DEL ENTRACHEQUE TRAE LOS DATOS DE LOS CHEQUES QUE SE HAYAN USADO'''

listche = "select * from entracheque where cliente='" & Label7 & "' and r_interno=" & Val(Label4) & ""
TABLA.Open listche, conexion_BD

'Call ALTAGRID_CHEQUES
li = 0
Call ALTAGRID_CHEQUES_CLIENTES

Do While Not TABLA.EOF

    li = li + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(li, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(li, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(li, 2) = TABLA!importe
    MSFlexGrid1.TextMatrix(li, 3) = TABLA!banco
    MSFlexGrid1.TextMatrix(li, 4) = TABLA!fecha_vto
    If TABLA!importe < "0" Then
        importe = CDbl(MSFlexGrid1.TextMatrix(li, 2)) * -1
        MSFlexGrid1.TextMatrix(li, 2) = importe
    End If '

    TABLA.MoveNext
Loop
TABLA.Close
        
''' COMO RETENCIONES NO HAY COMPLETAMOS EL REMITO''
Label17 = ""
reten = "select * from retencion where r_interno=" & Val(Label4) & ""
TABLA.Open reten, conexion_BD
        
    Do While Not TABLA.EOF
        Label17 = TABLA!importe
    TABLA.MoveNext
    Loop
        
TABLA.Close
End Select
End Sub


Private Sub Combo2_Click()
Call ALTAGRID

pers = "select * from remito_interno where destino='" & Combo2 & "' order by fecha"
TABLA.Open pers, conexion_BD

Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        If IsNull(TABLA!destino) Then
            MSFlexGrid2.TextMatrix(lin, 0) = ""
        Else
            MSFlexGrid2.TextMatrix(lin, 0) = TABLA!destino
        End If
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!fecha
        MSFlexGrid2.TextMatrix(lin, 2) = TABLA!n_remito
        If MSFlexGrid2.TextMatrix(lin, 2) = MSFlexGrid2.TextMatrix(lin - 1, 2) Then
            MSFlexGrid2.Row = lin
            MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
            lin = lin - 1
        Else
        
            MSFlexGrid2.TextMatrix(lin, 2) = TABLA!n_remito
    
        End If
    TABLA.MoveNext
Loop
TABLA.Close

Frame4.BorderStyle = 0
Frame4.Visible = True
End Sub

Private Sub Command1_Click()
Call IMPRIMIR

End Sub

Private Sub Command2_Click()
Frame1.Visible = False
Frame4.Visible = False
End Sub

Private Sub Command3_Click()
Select Case CTACTE
    Case "CLIENTES"
        
        cli = "select * from mov_clientes order by cliente"
        TABLA.Open cli, conexion_BD

        Do While Not TABLA.EOF
            lin = lin + 1
            MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
            MSFlexGrid2.TextMatrix(lin, 0) = TABLA!cliente
            MSFlexGrid2.TextMatrix(lin, 1) = TABLA!fecha
            MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
            
            If MSFlexGrid2.TextMatrix(lin, 2) = MSFlexGrid2.TextMatrix(lin - 1, 2) Then
                MSFlexGrid2.Row = lin
                MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
                lin = lin - 1
            Else
        
                MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
    
            End If
        TABLA.MoveNext
        Loop
        TABLA.Close
    Case "PROVEEDORES"
    
        Call ALTAGRID

        pro = "select * from mov_proveedor order by proveedor"
        TABLA.Open pro, conexion_BD

        Do While Not TABLA.EOF
            lin = lin + 1
            MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
            MSFlexGrid2.TextMatrix(lin, 0) = TABLA!proveedor
            MSFlexGrid2.TextMatrix(lin, 1) = TABLA!fecha
            MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
        
            If MSFlexGrid2.TextMatrix(lin, 2) = MSFlexGrid2.TextMatrix(lin - 1, 2) Then
                MSFlexGrid2.Row = lin
                MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
                lin = lin - 1
            Else
        
                MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
    
            End If
        TABLA.MoveNext
        Loop
        TABLA.Close
End Select
End Sub

Private Sub Command7_Click()
Frame4.Visible = False
End Sub

Private Sub Form_Load()
remi = "select * from remito_interno order by n_remito"
TABLA.Open remi, conexion_BD
Do While Not TABLA.EOF
    
    Combo1.AddItem TABLA!n_remito
    TABLA.MoveNext
Loop
TABLA.Close
End Sub

Private Sub ALTAGRID_CHEQUES()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 5
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 3
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "Interno Nº"
MSFlexGrid1.TextMatrix(0, 1) = "Cheque Nº"
MSFlexGrid1.TextMatrix(0, 2) = "Importe"
MSFlexGrid1.TextMatrix(0, 3) = "Banco"
MSFlexGrid1.TextMatrix(0, 4) = "Vencimiento"

MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.ColWidth(4) = 1500
End Sub
Private Sub ALTAGRID_CHEQUES_CLIENTES()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 5
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 3
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "Interno Nº"
MSFlexGrid1.TextMatrix(0, 1) = "Cheque Nº"
MSFlexGrid1.TextMatrix(0, 2) = "Importe"
MSFlexGrid1.TextMatrix(0, 3) = "Banco"
MSFlexGrid1.TextMatrix(0, 4) = "Vencimiento"

MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.ColWidth(4) = 1500
End Sub
Private Sub IMPRIMIR()

Printer.CurrentX = 30
Printer.CurrentY = 80
Printer.FontSize = 10
Printer.PaperSize = 9

Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Nº de recibo: "; Label4; Tab(55); "REIMPRESION"
Printer.Print Tab(110); Label5
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(10); " Recibí/mos de: "; Label7; " ."
Printer.Print Tab(10); ""
Printer.Print Tab(10); " la cantidad de pesos $ "; Label9
'Printer.Print Tab(10); ""
'Printer.Print Tab(10); " ( "; Label33; " )"
Printer.Print Tab(10); ""
Printer.Print Tab(10); " En concepto de: "; Label11;
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Según el siguiente detalle: Efectivo en $ "; Label15; Tab(70); "Retenciones: ", Label17
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Nº INTERNO"; Tab(25); "Nº CHEQUE"; Tab(43); "IMPORTE"; Tab(60); "BANCO"; Tab(83); "VENCIMIENTO"; Tab(120); '"FECHA VTO"; Tab(137); "FECHA MODIF"
Printer.Print Tab(10); "========================================================================="

For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    interno = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    cheque = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    importe = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    banco = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
    vto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4)
    
Printer.Print Tab(12); interno; Tab(27); cheque; Tab(45); importe; Tab(57); banco; Tab(85); vto; 'Tab(122); vto; Tab(139); depo
Next

Printer.Print Tab(10); "--------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); " Monto total en $ "; Label19
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(70); Label20
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -";

'''''''''''''''''''
'' IMPRIME COPIA ''
'''''''''''''''''''

Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Nº de recibo: "; Label4; Tab(55); "REIMPRESION"
Printer.Print Tab(110); Label5
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(10); " Recibí/mos de: "; Label7; " ."
Printer.Print Tab(10); ""
Printer.Print Tab(10); " la cantidad de pesos $ "; Label9
'Printer.Print Tab(10); ""
'Printer.Print Tab(10); " ( "; Label33; " )"
Printer.Print Tab(10); ""
Printer.Print Tab(10); " En concepto de: "; Label11;
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Según el siguiente detalle: Efectivo en $ "; Label15; Tab(70); "Retenciones: ", Label17
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Nº INTERNO"; Tab(25); "Nº CHEQUE"; Tab(43); "IMPORTE"; Tab(60); "BANCO"; Tab(83); "VENCIMIENTO"; Tab(120); '"FECHA VTO"; Tab(137); "FECHA MODIF"
Printer.Print Tab(10); "========================================================================="

For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    interno = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    cheque = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    importe = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    banco = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
    vto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4)
    
Printer.Print Tab(12); interno; Tab(27); cheque; Tab(45); importe; Tab(57); banco; Tab(85); vto; 'Tab(122); vto; Tab(139); depo
Next

Printer.Print Tab(10); "--------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); " Monto total en $ "; Label19
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(70); Label20
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -";
Printer.EndDoc
End Sub

Private Sub MSFlexGrid2_Click()


''' TRAE LOS DATOS PARA EL REMITO'''
Label20 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 0)
Label5 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 1)
Label4 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 2)
        
''' TRAE LOS DATOS DEL REMITO Y COMPLETA ESPACIOS
remito = "select * from remito_interno where n_remito= " & Val(Label4) & ""
TABLA.Open remito, conexion_BD

Label5 = TABLA!fecha
Label7 = TABLA!origen
Label9 = TABLA!total
If IsNull(TABLA!detalle) Then
    Label11 = ""
Else
    Label11 = TABLA!detalle
End If
Label19 = TABLA!total
Label20 = TABLA!destino
Label15 = "0"
Label17 = "0"
TABLA.Close
        
Select Case CTACTE


    Case "PERSONAL"
        
        cheques = "select * from mov_cheques where destino='" & Label20 & "' and fecha_mod='" & Label5 & "'"
        TABLA.Open cheques, conexion_BD
        Label15 = Label9
        Call ALTAGRID_CHEQUES
        Do While Not TABLA.EOF
            lin = lin + 1
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
            MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
            MSFlexGrid1.TextMatrix(lin, 2) = TABLA!importe
            'MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
            MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_vto
            If TABLA!importe < "0" Then
                importe = CDbl(MSFlexGrid1.TextMatrix(lin, 2)) * -1
                MSFlexGrid1.TextMatrix(lin, 2) = importe
            End If

            TABLA.MoveNext
        Loop
        TABLA.Close
        

        
    Case "CLIENTES"

        ''' DEL MOV_CLIENETS TRAE EL PAGO EN EFECTIVO'''
        cliente = "select * from mov_clientes where r_interno='" & Label4 & "'"
        TABLA.Open cliente, conexion_BD
        Do While Not TABLA.EOF
            If TABLA!n_cheque = "0" Then
                Label15 = TABLA!pago
            End If
            TABLA.MoveNext
        Loop
        TABLA.Close
        ''' DEL ENTRACHEQUE TRAE LOS DATOS DE LOS CHEQUES QUE SE HAYAN USADO'''
        listche = "select * from entracheque where cliente='" & Label7 & "' and r_interno=" & Val(Label4) & ""
        TABLA.Open listche, conexion_BD
        
        Call ALTAGRID_CHEQUES_CLIENTES
        Do While Not TABLA.EOF
            li = li + 1
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(li, 0) = TABLA!n_interno
            MSFlexGrid1.TextMatrix(li, 1) = TABLA!n_cheque
            MSFlexGrid1.TextMatrix(li, 2) = TABLA!importe
            MSFlexGrid1.TextMatrix(li, 3) = TABLA!banco
            MSFlexGrid1.TextMatrix(li, 4) = TABLA!fecha_vto
            If TABLA!importe < "0" Then
                importe = CDbl(MSFlexGrid1.TextMatrix(li, 2)) * -1
                MSFlexGrid1.TextMatrix(li, 2) = importe
            End If

            TABLA.MoveNext
        Loop
        TABLA.Close
        
        ''' COMO RETENCIONES NO HAY COMPLETAMOS EL REMITO''
        Label17 = ""
    
    Case "PROVEEDORES"
    
        ''' DEL MOV_PROVEEDOR TRAE EL PAGO EN EFECTIVO'''
        cliente = "select * from mov_proveedor where r_interno='" & Label4 & "'"
        TABLA.Open cliente, conexion_BD
        Do While Not TABLA.EOF
            If TABLA!n_cheque = "0" Then
                Label15 = TABLA!pago
            End If
            TABLA.MoveNext
        Loop
        TABLA.Close
        ''' DEL SALECHEQUE TRAE LOS DATOS DE LOS CHEQUES QUE SE HAYAN USADO'''

        listche = "select * from mov_cheques where fecha_mod='" & Label5 & "' and destino ='" & Label20 & "'"
        TABLA.Open listche, conexion_BD
        
        Call ALTAGRID_CHEQUES
        Do While Not TABLA.EOF
            lin = lin + 1
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
            MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
            MSFlexGrid1.TextMatrix(lin, 2) = TABLA!importe
            'MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
            MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_vto
            If TABLA!importe < "0" Then
                importe = CDbl(MSFlexGrid1.TextMatrix(lin, 2)) * -1
                MSFlexGrid1.TextMatrix(lin, 2) = importe
            End If

            TABLA.MoveNext
        Loop
        TABLA.Close
        
        ''' RETENCIONES ''
        reten = "select * from retencion where r_interno=" & Val(Label4) & ""
        TABLA.Open reten, conexion_BD
        
        Do While Not TABLA.EOF
            Label17 = TABLA!importe
        TABLA.MoveNext
        Loop
        
        TABLA.Close
End Select
Frame4.Visible = False
Frame1.Visible = True
End Sub

Private Sub Option1_Click()
If Option1 = True Then
    Frame3.Visible = True
    Frame4.Visible = False
    Frame1.Visible = False
    Label2.Visible = False
    Combo1.Visible = False
Else
    Frame3.Visible = False
    Frame4.Visible = False
    Frame1.Visible = False
    Label2.Visible = True
    Combo1.Visible = True
End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
    Frame3.Visible = False
    Frame4.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
    Label2.Visible = True
    Combo1.Visible = True
Else
    Frame3.Visible = True
    Frame4.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
    Label2.Visible = False
    Combo1.Visible = False
    
End If
End Sub

Private Sub Option3_Click()
CTACTE = "CLIENTES"
'''CLIENTES'''
Call ALTAGRID

cli = "select * from mov_clientes order by cliente"
TABLA.Open cli, conexion_BD

Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!cliente
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!fecha
        MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
        If MSFlexGrid2.TextMatrix(lin, 2) = MSFlexGrid2.TextMatrix(lin - 1, 2) Then
            MSFlexGrid2.Row = lin
            MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
            lin = lin - 1
        Else
        
            MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
    
        End If
    TABLA.MoveNext
Loop
TABLA.Close

Frame4.BorderStyle = 0
Frame4.Visible = True
Frame5.Visible = False
Option3 = False

End Sub

Private Sub Option4_Click()
Option4 = False

CTACTE = "PERSONAL"

rrhh = "select * from alta_rrhh order by nombre_rrhh"
TABLA.Open rrhh, conexion_BD

Do While Not TABLA.EOF
    Combo2.AddItem TABLA!nombre_rrhh
    TABLA.MoveNext
Loop
Frame5.BorderStyle = 0
Frame5.Visible = True
Frame4.Visible = False
TABLA.Close


End Sub

Private Sub Option5_Click()
Option5 = False
'''PROVEEDORES'''
CTACTE = "PROVEEDORES"
Call ALTAGRID

pro = "select * from mov_proveedor order by proveedor"
TABLA.Open pro, conexion_BD

Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!proveedor
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!fecha
        MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
        If MSFlexGrid2.TextMatrix(lin, 2) = MSFlexGrid2.TextMatrix(lin - 1, 2) Then
            MSFlexGrid2.Row = lin
            MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
            lin = lin - 1
        Else
        
            MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
    
        End If
    TABLA.MoveNext
Loop
TABLA.Close

Frame4.BorderStyle = 0
Frame4.Visible = True
Frame5.Visible = False

End Sub
Private Sub ALTAGRID()
With MSFlexGrid2
.FixedCols = 0
.Cols = 3
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "Nombre/Razón social"
.TextMatrix(0, 1) = "Fecha"
.TextMatrix(0, 2) = "Remito"

.ColWidth(0) = 2700
.ColWidth(1) = 1500
.ColWidth(2) = 1000
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case CTACTE

        Case "CLIENTES"
            Q = "select * from mov_clientes order by cliente"
            TABLA.Open Q, conexion_BD
            sal = 0
            lin = 0
            MSFlexGrid2.Clear
            Call ALTAGRID
    
            Do While Not TABLA.EOF
    
            If UCase(Left(TABLA!cliente, Len(Text1))) = UCase(Text1) Then
                lin = lin + 1
                MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
                MSFlexGrid2.TextMatrix(lin, 0) = TABLA!cliente
                MSFlexGrid2.TextMatrix(lin, 1) = TABLA!fecha
                MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
                
                If MSFlexGrid2.TextMatrix(lin, 2) = MSFlexGrid2.TextMatrix(lin - 1, 2) Then
                    MSFlexGrid2.Row = lin
                    MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
                    lin = lin - 1
                Else
        
                    MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
    
            End If
        End If
        TABLA.MoveNext
    Loop
    TABLA.Close

Case "PROVEEDORES"

    Call ALTAGRID

    pro = "select * from mov_proveedor order by proveedor"
    TABLA.Open pro, conexion_BD

    Do While Not TABLA.EOF
        If UCase(Left(TABLA!proveedor, Len(Text1))) = UCase(Text1) Then
            lin = lin + 1
            MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
            MSFlexGrid2.TextMatrix(lin, 0) = TABLA!proveedor
            MSFlexGrid2.TextMatrix(lin, 1) = TABLA!fecha
            MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
        
            If MSFlexGrid2.TextMatrix(lin, 2) = MSFlexGrid2.TextMatrix(lin - 1, 2) Then
                MSFlexGrid2.Row = lin
                MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
                lin = lin - 1
            Else
        
                MSFlexGrid2.TextMatrix(lin, 2) = TABLA!r_interno
    
            End If
        End If
        TABLA.MoveNext
    Loop
    TABLA.Close
End Select
End If
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form Listado_x_ctacte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados de Saldos de Cuentas Corrientes"
   ClientHeight    =   9705
   ClientLeft      =   1455
   ClientTop       =   960
   ClientWidth     =   12405
   Icon            =   "Listado_x_ctacte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   12405
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir selección"
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
      Left            =   10800
      Picture         =   "Listado_x_ctacte.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cargando. Por favor espere."
      Height          =   615
      Left            =   4080
      TabIndex        =   36
      Top             =   9000
      Visible         =   0   'False
      Width           =   3735
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   200
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listar Todo"
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
      Left            =   10800
      Picture         =   "Listado_x_ctacte.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Movimiento"
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
      Left            =   10800
      Picture         =   "Listado_x_ctacte.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Atrás"
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
      Left            =   10800
      Picture         =   "Listado_x_ctacte.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   10800
      Picture         =   "Listado_x_ctacte.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   9375
      Begin VB.Frame Frame4 
         Height          =   6015
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   9375
         Begin VB.PictureBox picUnchecked 
            Height          =   285
            Left            =   8760
            Picture         =   "Listado_x_ctacte.frx":213C
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   42
            Top             =   600
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.PictureBox picChecked 
            Height          =   285
            Left            =   8760
            Picture         =   "Listado_x_ctacte.frx":247E
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Buscar"
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
            Picture         =   "Listado_x_ctacte.frx":27C0
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   240
            Width           =   1095
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
            Left            =   3240
            TabIndex        =   39
            Top             =   360
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
            Height          =   4095
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Visible         =   0   'False
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   7223
            _Version        =   393216
            BackColor       =   16777152
            BackColorBkg    =   -2147483633
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
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Label18"
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
            Left            =   6720
            TabIndex        =   45
            Top             =   5400
            Width           =   2295
         End
         Begin VB.Label Label17 
            Caption         =   "Total saldo:"
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
            TabIndex        =   44
            Top             =   5400
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Filtrar por saldo inferior a:  $"
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
            TabIndex        =   38
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Atrás"
         Height          =   615
         Left            =   7800
         Picture         =   "Listado_x_ctacte.frx":2D4A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Atrás"
         Height          =   615
         Left            =   7800
         Picture         =   "Listado_x_ctacte.frx":32D4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Atrás"
         Height          =   615
         Left            =   7800
         Picture         =   "Listado_x_ctacte.frx":385E
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
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
         Left            =   4800
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   5415
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   9551
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   5535
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   9763
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5535
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   9763
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
      Begin VB.Label Label13 
         Caption         =   "<---- aca hay botones escondidos"
         Height          =   855
         Left            =   9000
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Buscar por Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   24
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.Line Line8 
         X1              =   9120
         X2              =   9120
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   120
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   9120
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   9120
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line4 
         X1              =   7440
         X2              =   7440
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line3 
         X1              =   5640
         X2              =   5640
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line2 
         X1              =   3120
         X2              =   3120
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9120
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label3 
         Caption         =   "Pagos Realizados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Elegir una Opción:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
      Begin VB.OptionButton OptionProve 
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
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton OptionPersonal 
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
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptionClientes 
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
         Height          =   240
         Left            =   600
         TabIndex        =   3
         Top             =   735
         Width           =   1095
      End
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Height          =   375
      Left            =   2400
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "L14= opcion"
      Height          =   375
      Left            =   360
      TabIndex        =   30
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Buscar por Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   28
      Top             =   2880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label10 
      Caption         =   "L10= NOMBRE"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Saldos de Cuentas Corrientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "Listado_x_ctacte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FILAS As String
Dim strChecked As String


Private Sub Command1_Click()
Call IMPRIMIR
End Sub
Private Sub IMPRIMIR()
Printer.Orientation = 1
Printer.FontSize = 11
Printer.Font = arial
Printer.FontBold = False
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Fecha: "; Date; Tab(49); SISTEMA
Printer.Print Tab(40); SISTEMA_DIR
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "SALDOS DE CUENTAS CORRIENTES "; Label14
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.FontSize = 9
If Label14 = "PERSONAL" Then
    Printer.Print Tab(13); "NOMBRE/RAZON SOCIAL"; Tab(65); "ADELANTO"; Tab(85); "PREMIO"; Tab(103); "SUELDO"; Tab(120); "SALDO"; ' Tab(125); "FECHA VTO"
    Printer.Print Tab(10); "==============================================================================================="
Else

    Printer.Print Tab(13); "NOMBRE/RAZON SOCIAL"; Tab(65); "PAGO"; Tab(85); "FACTURA"; Tab(103); "RETEN"; Tab(120); "SALDO"; ' Tab(125); "FECHA VTO"
    Printer.Print Tab(10); "============================================================================================"
End If
With MSFlexGrid5
For i = 1 To .Rows - 1
    .Row = i
    cliente = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 1)
    pago = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 2)
    factura = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 3)
    RETENCION = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 4)
    saldo = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 5)


Printer.Print Tab(10); cliente; Tab(65); pago; Tab(85); factura; Tab(103); RETENCION; Tab(120); saldo; 'Tab(127); vto; 'Tab(130); modif
Printer.Print Tab(10); "-------------------------------------------------------------------------------------------------------------------------------------------------------------------"


    If i = 32 Or i = 68 Or i = 103 Or i = 138 Then
        Printer.NewPage
        Printer.Print Tab(10); ""
        Printer.Print Tab(10); ""
        Printer.Print Tab(10); ""
        
        If Label14 = "PERSONAL" Then
            Printer.Print Tab(13); "NOMBRE/RAZON SOCIAL"; Tab(65); "ADELANTO"; Tab(85); "PREMIO"; Tab(103); "SUELDO"; Tab(120); "SALDO"; ' Tab(125); "FECHA VTO"
            Printer.Print Tab(10); "==============================================================================================="
        Else

            Printer.Print Tab(13); "NOMBRE/RAZON SOCIAL"; Tab(65); "PAGO"; Tab(85); "FACTURA"; Tab(103); "RETEN"; Tab(120); "SALDO"; ' Tab(125); "FECHA VTO"
            Printer.Print Tab(10); "============================================================================================"
        End If
    End If
'End If
Next

Printer.Print Tab(85); Label17; "  "; Label18
End With

Printer.EndDoc
End Sub
Private Sub Command2_Click()
Frame4.Visible = False
Command2.Visible = False
Command1.Visible = False
End Sub

Private Sub Command3_Click()
datos = Label10
sValor = Label14

If sValor = "PERSONAL" Then
    datos1 = Label6
    mov_jornales.Show
    
Else
    Detalles_ctas.Show
End If
End Sub

Private Sub Command4_Click()
MSFlexGrid5.Visible = False


Frame4.Visible = True
Frame5.Visible = True
Select Case Label14
Case "CLIENTES"
    Call ALTALISTA4
    ''' LLAMA A LA RUTINA PARA CARGAR LOS CLIENTES EN EL FLEXGRID5
    Call CARGA_CLIENTES
    
Case "PROVEEDORES"
    Call ALTALISTA4
    ''' LLAMA A LA RUTINA PARA CARGAR LOS PROVEEDORES EN EL FLEXGRID5
    Call CARGA_PROVEEDORES
    
Case "PERSONAL"
    Call ALTALISTA5
    Call CARGA_PERSONAL
End Select

Command2.Visible = True
Command1.Visible = True
MSFlexGrid5.Visible = True
Frame5.Visible = False


With MSFlexGrid5
TOTFILA = FILAS - 2
total = 0
For K = 1 To TOTFILA

    If .TextMatrix(K, 5) <> "" Then
        total = CDbl(total) + .TextMatrix(K, 5)

    Else
        .Row = K
        .RemoveItem (.Row)
        'k = k - 1
        K = TOTFILA '+ 1
        alto = TOTFILA

    End If
Next

Label18 = Format(total, "currency")
End With


End Sub


Private Sub CARGA_CLIENTES()
For i = 0 To MSFlexGrid2.Rows - 1
    SQL = "select *  from mov_clientes where cliente='" & MSFlexGrid2.TextMatrix(i, 1) & "'"
    TABLA.Open SQL, conexion_BD
    
    With MSFlexGrid5
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 1) = TABLA!cliente

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 1) = .TextMatrix(lin - 1, 1) Then
            
            If IsNull(TABLA!pago) Then
                cta1 = 0
                .TextMatrix(lin, 2) = Format(cta1, "currency")
                cta1 = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
            Else
                .TextMatrix(lin, 2) = Format(TABLA!pago, "currency")
                cta1 = CDbl(.TextMatrix(lin - 1, 2)) + .TextMatrix(lin, 2)
            
            End If

            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 2) = Format(cta1, "currency")
            .TextMatrix(lin, 3) = Format(0, "currency")
            ProgressBar1.value = lin
            
            Else
 
                '.TextMatrix(lin, 1) = TABLA!cliente
                
                If IsNull(TABLA!pago) Then
                    .TextMatrix(lin, 2) = Format("0", "currency")
                Else
                    .TextMatrix(lin, 2) = Format(TABLA!pago, "currency")
                End If

                cta1 = 0
                cta2 = 0
                .TextMatrix(lin, 3) = Format(cta2, "currency")
                
                ProgressBar1.value = lin

                
            End If

        TABLA.MoveNext
    Loop
    TABLA.Close

   End With
   
    SQL = "select *  from clientes_a_pagar where cliente='" & MSFlexGrid2.TextMatrix(i, 1) & "'"
    TABLA.Open SQL, conexion_BD
    

    With MSFlexGrid5

    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 1) = TABLA!cliente

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 1) = .TextMatrix(lin - 1, 1) Then
            
            If IsNull(TABLA!monto) Then
                cta2 = 0
                .TextMatrix(lin, 3) = Format(cta2, "currency")
                cta2 = CDbl(.TextMatrix(lin, 3)) + CDbl(.TextMatrix(lin - 1, 3))
            Else
                .TextMatrix(lin, 3) = TABLA!monto
                cta2 = CDbl(.TextMatrix(lin - 1, 3)) + TABLA!monto
            
            End If

            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 3) = Format(cta2, "currency")
            ProgressBar1.value = lin
            
            Else
 
                .TextMatrix(lin, 1) = TABLA!cliente
                
                If IsNull(TABLA!monto) Then
                    .TextMatrix(lin, 3) = Format("0", "currency")
                Else
                    .TextMatrix(lin, 3) = Format(TABLA!monto, "currency")
                End If

                cta1 = 0
                cta2 = 0
                If .TextMatrix(lin, 2) = "" Then
                    .TextMatrix(lin, 2) = Format(0, "currency")
                End If
                
                ProgressBar1.value = lin
            End If

        TABLA.MoveNext
    Loop
    TABLA.Close

   End With
 Next
'''' SACAMOS EL SALDO
With MSFlexGrid5
'ProgressBar1.Min = 0
'ProgressBar1.Max = i

For i = 1 To .Rows - 1
    Set .CellPicture = picUnchecked.Picture
    If .TextMatrix(i, 2) = "" Then
        .TextMatrix(i, 2) = 0
    End If
    If .TextMatrix(i, 3) = "" Then
        .TextMatrix(i, 3) = 0
    End If
    saldo = CDbl(.TextMatrix(i, 2)) - CDbl(.TextMatrix(i, 3))
    .TextMatrix(i, 5) = Format(saldo, "currency")
    FILAS = i
Next

For lin = 1 To .Rows - 1
    '.Row = .Rows + 1
    If .TextMatrix(lin, 5) = "$ 0,00" Then
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        .Rows = .Rows + 1
        FILAS = .Rows

   End If
'    ProgressBar1.value = i
Next

End With

End Sub
Private Sub CARGA_PROVEEDORES()
For i = 1 To MSFlexGrid4.Rows - 1
    SQL = "select *  from mov_proveedor where proveedor='" & MSFlexGrid4.TextMatrix(i, 1) & "'"
    TABLA.Open SQL, conexion_BD

    With MSFlexGrid5
 
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 1) = TABLA!proveedor

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 1) = .TextMatrix(lin - 1, 1) Then
            
            If IsNull(TABLA!pago) Then
                cta1 = 0
                .TextMatrix(lin, 2) = Format(cta1, "currency")
                cta1 = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
            Else
                .TextMatrix(lin, 2) = Format(TABLA!pago, "currency")
                cta1 = CDbl(.TextMatrix(lin - 1, 2)) + TABLA!pago
            
            End If

            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 2) = Format(cta1, "currency")
            .TextMatrix(lin, 3) = Format(0, "currency")
            .TextMatrix(lin, 4) = Format(0, "currency")
            
            Else
 
                .TextMatrix(lin, 1) = TABLA!proveedor
                
                If IsNull(TABLA!pago) Then
                    .TextMatrix(lin, 2) = Format("0", "currency")
                Else
                    .TextMatrix(lin, 2) = Format(TABLA!pago, "currency")
                End If

                cta1 = 0
                cta2 = 0
                cta3 = 0
                .TextMatrix(lin, 3) = Format(cta2, "currency")
                .TextMatrix(lin, 4) = Format(cta3, "currency")
                
            End If
            ProgressBar1.value = lin
            
        TABLA.MoveNext
    Loop
    TABLA.Close

   End With
   
    SQL = "select *  from prove_a_pagar where proveedor='" & MSFlexGrid4.TextMatrix(i, 1) & "'"
    TABLA.Open SQL, conexion_BD
    
    With MSFlexGrid5
    
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 1) = TABLA!proveedor

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 1) = .TextMatrix(lin - 1, 1) Then
            
            If IsNull(TABLA!monto) Then
                cta2 = 0
                .TextMatrix(lin, 3) = Format(cta2, "currency")
                cta2 = CDbl(.TextMatrix(lin, 3)) + CDbl(.TextMatrix(lin - 1, 3))
            Else
                .TextMatrix(lin, 3) = TABLA!monto
                cta2 = CDbl(.TextMatrix(lin - 1, 3)) + TABLA!monto
            
            End If

            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 3) = Format(cta2, "currency")
            .TextMatrix(lin, 4) = Format(0, "currency")
            
            Else
 
                .TextMatrix(lin, 1) = TABLA!proveedor
                
                If IsNull(TABLA!monto) Then
                    .TextMatrix(lin, 3) = Format(0, "currency")
                Else
                    .TextMatrix(lin, 3) = Format(TABLA!monto, "currency")
                End If

                cta1 = 0
                cta2 = 0
                If .TextMatrix(lin, 2) = "" Then
                    .TextMatrix(lin, 2) = Format(0, "currency")
                End If
                .TextMatrix(lin, 4) = Format(0, "currency")
                
            End If
            ProgressBar1.value = lin
        TABLA.MoveNext
    Loop
    TABLA.Close

   End With
   
   ''' RETENCIONES
   
     SQL = "select *  from retencion where destino='" & MSFlexGrid4.TextMatrix(i, 1) & "'"
    TABLA.Open SQL, conexion_BD
    
    With MSFlexGrid5
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 1) = TABLA!destino

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 1) = .TextMatrix(lin - 1, 1) Then
            
            If IsNull(TABLA!importe) Then
                cta3 = 0
                .TextMatrix(lin, 4) = Format(cta3, "currency")
                cta3 = CDbl(.TextMatrix(lin, 4)) + CDbl(.TextMatrix(lin - 1, 4))
            Else
                .TextMatrix(lin, 4) = Format(TABLA!importe, "currency")
                cta3 = CDbl(.TextMatrix(lin - 1, 4)) + TABLA!importe
            
            End If

            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 4) = Format(cta3, "currency")
            
            Else
 
                .TextMatrix(lin, 1) = TABLA!destino
                
                If IsNull(TABLA!importe) Then
                    .TextMatrix(lin, 4) = Format(0, "currency")
                Else
                    .TextMatrix(lin, 4) = Format(TABLA!importe, "currency")
                End If
                cta3 = 0
               
       
            End If

        TABLA.MoveNext
    Loop
    TABLA.Close

   End With
 
 
 Next
 
'''' SACAMOS EL SALDO
With MSFlexGrid5
'ProgressBar1.Min = 0
'ProgressBar1.Max = i
For i = 1 To .Rows - 1
    Set .CellPicture = picUnchecked.Picture
    If .TextMatrix(i, 2) = "" Then
        .TextMatrix(i, 2) = 0
    End If
    If .TextMatrix(i, 3) = "" Then
        .TextMatrix(i, 3) = 0
    End If
    If .TextMatrix(i, 4) = "" Then
        .TextMatrix(i, 4) = 0
    End If
        
    pagoyret = CDbl(.TextMatrix(i, 2)) + CDbl(.TextMatrix(i, 4))
    suma = Format(pagoyret, "currency")
    saldo = 0
    saldo = CDbl(suma) - CDbl(.TextMatrix(i, 3))
    .TextMatrix(i, 5) = Format(saldo, "currency")
    
Next

For i = 0 To .Rows - 1
    '.Row = .Rows + 1
    If .TextMatrix(i, 5) = "$ 0,00" Then
        .Row = i
        .RemoveItem (.Row)
        i = i - 1
        .Rows = .Rows + 1
        FILAS = .Rows
    End If
   ' ProgressBar1.value = i
Next
End With
End Sub
Private Sub CARGA_PERSONAL()
For i = 1 To MSFlexGrid1.Rows - 1
    SQL = "select *  from mov_rrhh where nombre_rrhh='" & MSFlexGrid1.TextMatrix(i, 1) & "'"
    TABLA.Open SQL, conexion_BD
    ''' ADELANTOS
    With MSFlexGrid5
 
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 1) = TABLA!nombre_rrhh

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 1) = .TextMatrix(lin - 1, 1) Then
            
            If IsNull(TABLA!adelanto) Then
                cta1 = 0
                .TextMatrix(lin, 2) = Format(cta1, "currency")
                cta1 = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
            Else
                .TextMatrix(lin, 2) = TABLA!adelanto
                cta1 = CDbl(.TextMatrix(lin - 1, 2)) + TABLA!adelanto
            
            End If

            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 2) = Format(cta1, "currency")
            .TextMatrix(lin, 3) = 0
            .TextMatrix(lin, 4) = 0
            
            Else
 
                .TextMatrix(lin, 1) = TABLA!nombre_rrhh
                
                If IsNull(TABLA!adelanto) Then
                    .TextMatrix(lin, 2) = Format(0, "currency")
                Else
                    .TextMatrix(lin, 2) = Format(TABLA!adelanto, "currency")
                End If

                cta1 = 0
                cta2 = 0
                cta3 = 0
                .TextMatrix(lin, 3) = Format(cta2, "currency")
                .TextMatrix(lin, 4) = Format(cta3, "currency")
                
            End If
            ProgressBar1.value = lin
            
        TABLA.MoveNext
    Loop
    TABLA.Close

   End With
   
    SQL = "select *  from mov_rrhh where nombre_rrhh='" & MSFlexGrid1.TextMatrix(i, 1) & "'"
    TABLA.Open SQL, conexion_BD
    ''' PREMIO
    With MSFlexGrid5
    
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 1) = TABLA!nombre_rrhh

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 1) = .TextMatrix(lin - 1, 1) Then
            
            If IsNull(TABLA!premio) Then
                cta3 = 0
                .TextMatrix(lin, 3) = Format(cta3, "currency")
                cta3 = CDbl(.TextMatrix(lin, 3)) + CDbl(.TextMatrix(lin - 1, 3))
            Else
                .TextMatrix(lin, 3) = TABLA!premio
                cta3 = CDbl(.TextMatrix(lin - 1, 3)) + TABLA!premio
            
            End If

            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 3) = Format(cta3, "currency")
            .TextMatrix(lin, 4) = 0
            
            Else
 
                .TextMatrix(lin, 1) = TABLA!nombre_rrhh
                
                If IsNull(TABLA!premio) Then
                    cta3 = 0
                    .TextMatrix(lin, 3) = Format(0, "currency")
                Else
                    .TextMatrix(lin, 3) = Format(TABLA!premio, "currency")
                    cta3 = .TextMatrix(lin, 3)
                    
                End If

                cta1 = 0
                cta2 = 0
                
                If .TextMatrix(lin, 2) = "" Then
                    .TextMatrix(lin, 2) = Format(0, "currency")
                End If
                .TextMatrix(lin, 4) = Format(0, "currency")
                
            End If
            ProgressBar1.value = lin
        TABLA.MoveNext
    Loop
    TABLA.Close

   End With
   
   ''' SUELDO
   
     SQL = "select *  from mov_rrhh where nombre_rrhh='" & MSFlexGrid1.TextMatrix(i, 1) & "'"
    TABLA.Open SQL, conexion_BD
    
    With MSFlexGrid5
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 1) = TABLA!nombre_rrhh

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 1) = .TextMatrix(lin - 1, 1) Then
            
            If IsNull(TABLA!total_pesos) Then
                cta4 = 0
                .TextMatrix(lin, 4) = Format(cta4, "currency")
                cta4 = CDbl(.TextMatrix(lin, 4)) + CDbl(.TextMatrix(lin - 1, 4))
            Else
                .TextMatrix(lin, 4) = TABLA!total_pesos
                cta4 = CDbl(.TextMatrix(lin - 1, 4)) + TABLA!total_pesos
            
            End If

            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 4) = Format(cta4, "currency")
            
            Else
 
                .TextMatrix(lin, 1) = TABLA!total_pesos
                
                If IsNull(TABLA!importe) Then
                    cta4 = "0"
                Else
                    .TextMatrix(lin, 4) = Format(TABLA!total_pesos, "currency")
                    cta4 = .TextMatrix(lin, 4)
                End If
                cta4 = 0
               
       
            End If

        TABLA.MoveNext
    Loop
    TABLA.Close

   End With
 
 
 Next
 
'''' SACAMOS EL SALDO
With MSFlexGrid5
'ProgressBar1.Min = 0
'ProgressBar1.Max = i
For i = 1 To .Rows - 1
    Set .CellPicture = picUnchecked.Picture
    If .TextMatrix(i, 2) = "" Then
        .TextMatrix(i, 2) = 0
    End If
    If .TextMatrix(i, 3) = "" Then
        .TextMatrix(i, 3) = 0
    End If
    If .TextMatrix(i, 4) = "" Then
       .TextMatrix(i, 4) = 0
    End If

    adelypre = CDbl(.TextMatrix(i, 2)) + CDbl(.TextMatrix(i, 3))
    suma = Format(adelypre, "currency")
    saldo = 0
    saldo = CDbl(suma) - CDbl(.TextMatrix(i, 4))
    .TextMatrix(i, 5) = Format(saldo, "currency")
Next

For i = 0 To .Rows - 1
    '.Row = .Rows + 1
    If .TextMatrix(i, 5) = "$ 0,00" Then
        .Row = i
        .RemoveItem (.Row)
        i = i - 1
        .Rows = .Rows + 1
        FILAS = .Rows
    End If
    ProgressBar1.value = i
Next
End With

End Sub
Private Sub Command5_Click()
MSFlexGrid1.Clear
MSFlexGrid1.Visible = True
MSFlexGrid2.Visible = False
MSFlexGrid4.Visible = False
Text1.Visible = True
Label6 = ""
Label7 = ""
Label8 = ""
Label9 = ""
Label10 = ""
Call ALTALISTA
W = "select * from alta_rrhh order by nombre_rrhh"
TABLA.Open W, conexion_BD
Frame2.Visible = True
Frame3.Visible = True
    lin = 0
    Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!id_rrhh
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!nombre_rrhh
        TABLA.MoveNext
    Loop
    
    TABLA.Close
End Sub

Private Sub Command6_Click()
MSFlexGrid4.Clear
MSFlexGrid1.Visible = False
MSFlexGrid2.Visible = False
MSFlexGrid4.Visible = True
Text1.Visible = False
Text2.Visible = True
Command5.Visible = False
Command6.Visible = True

Label6 = ""
Label7 = ""
Label8 = ""
Label9 = ""
Label10 = ""

Call ALTALISTA3
W = "select * from proveedores order by nombre_prove asc"
TABLA.Open W, conexion_BD

Frame3.Visible = True
    lin = 0
    Do While Not TABLA.EOF
    lin = lin + 1
        MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
        MSFlexGrid4.TextMatrix(lin, 0) = TABLA!id_prove
        MSFlexGrid4.TextMatrix(lin, 1) = TABLA!nombre_prove
        TABLA.MoveNext
    Loop
    
    TABLA.Close
End Sub

Private Sub Command7_Click()
MSFlexGrid2.Clear
Call ALTALISTA2

com = "select * from clientes order by nombre_cliente"
TABLA.Open com, conexion_BD
Frame3.Visible = True

Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
    
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!id_cliente
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!nombre_cliente

        TABLA.MoveNext
    Loop
    
    TABLA.Close

End Sub

Private Sub Command8_Click()

With MSFlexGrid5

For a = 1 To .Rows - 1

    If .TextMatrix(a, 5) = "SALDO" Or .TextMatrix(a, 5) = "" Then
        

    Else
        If .TextMatrix(a, 5) >= Val(Text4) Then
            .Row = a
            .RemoveItem (.Row)
            a = a - 1
            .Rows = .Rows + 1
            FILAS = .Rows
        End If
    End If

Next

End With

MSFlexGrid5.TextMatrix(0, 5) = "SALDO"
'MSFlexGrid5.TextMatrix(a - 1, 5) = ""

With MSFlexGrid5
TOTFILA = FILAS - 1
total = 0
For i = 1 To TOTFILA

    'If .TextMatrix(i, 5) = "" Then
    '        .Row = i
    '        .RemoveItem (.Row)
    '        i = i - 1
    'Else
    '    total = CDbl(total) + .TextMatrix(i, 5)
    'End If
    If .TextMatrix(i, 5) <> "" Then
        total = CDbl(total) + .TextMatrix(i, 5)
    'Else
    '    .Row = i
    '    .RemoveItem (.Row)
        
    End If
Next
Label18 = Format(total, "currency")
End With
End Sub

Private Sub Command9_Click()
'imprimir solo la seleccion del grid 5

Printer.Orientation = 1
Printer.FontSize = 11
Printer.Font = arial
Printer.FontBold = False
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Fecha: "; Date; Tab(49); SISTEMA
Printer.Print Tab(40); SISTEMA_DIR
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "SALDOS DE CUENTAS CORRIENTES "; Label14
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.FontSize = 9
If Label14 = "PERSONAL" Then
    Printer.Print Tab(13); "NOMBRE/RAZON SOCIAL"; Tab(65); "ADELANTO"; Tab(85); "PREMIO"; Tab(103); "SUELDO"; Tab(120); "SALDO"; ' Tab(125); "FECHA VTO"
    Printer.Print Tab(10); "==============================================================================================="
Else

    Printer.Print Tab(13); "NOMBRE/RAZON SOCIAL"; Tab(65); "PAGO"; Tab(85); "FACTURA"; Tab(103); "RETEN"; Tab(120); "SALDO"; ' Tab(125); "FECHA VTO"
    Printer.Print Tab(10); "============================================================================================"
End If
With MSFlexGrid5

For i = 1 To .Row
    .Row = i
    If .TextMatrix(.Row, 6) = "1" Then
    
    cliente = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 1)
    pago = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 2)
    factura = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 3)
    RETENCION = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 4)
    saldo = Me.MSFlexGrid5.TextMatrix(Me.MSFlexGrid5.Row, 5)


Printer.Print Tab(10); cliente; Tab(65); pago; Tab(85); factura; Tab(103); RETENCION; Tab(120); saldo; 'Tab(127); vto; 'Tab(130); modif
Printer.Print Tab(10); "-------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    If i = 32 Or i = 68 Or i = 103 Or i = 138 Then
        Printer.NewPage
        Printer.Print Tab(10); ""
        Printer.Print Tab(10); ""
        Printer.Print Tab(10); ""
        
        If Label14 = "PERSONAL" Then
            Printer.Print Tab(13); "NOMBRE/RAZON SOCIAL"; Tab(65); "ADELANTO"; Tab(85); "PREMIO"; Tab(103); "SUELDO"; Tab(120); "SALDO"; ' Tab(125); "FECHA VTO"
            Printer.Print Tab(10); "==============================================================================================="
        Else

            Printer.Print Tab(13); "NOMBRE/RAZON SOCIAL"; Tab(65); "PAGO"; Tab(85); "FACTURA"; Tab(103); "RETEN"; Tab(120); "SALDO"; ' Tab(125); "FECHA VTO"
            Printer.Print Tab(10); "============================================================================================"
        End If
    End If
End If
Next

Printer.Print Tab(85); Label17; "  "; Label18
End With

Printer.EndDoc
End Sub

Private Sub Form_Load()
ProgressBar1.Max = 1000
ProgressBar1.Min = 0
Call ALTALISTA
End Sub



Private Sub MSFlexGrid1_Click()
' personal
    Label10 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    
    Q = "select * from mov_rrhh where id_rrhh= " & Val(Label10) & ""
    TABLA.Open Q, conexion_BD
    sal = 0
    Do While Not TABLA.EOF = True
        If IsNull(TABLA!premio) Then
            premio = 0 + CDbl(premio)
        Else
            premio = CDbl(premio) + TABLA!premio
        End If
        
        sueldo = CDbl(sueldo) + TABLA!total_pesos
        adelanto = CDbl(adelanto) + TABLA!adelanto
        TABLA.MoveNext
        
    Loop
        suma = CDbl(sueldo) + CDbl(premio)
        saldo = CDbl(suma) - CDbl(adelanto)

        
    Label9 = Format(saldo, "currency")
    Label6 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    TABLA.Close
    Frame2.Visible = True
    MSFlexGrid2.Visible = False


End Sub

Private Sub MSFlexGrid2_Click()
'CLIENTES
    Label10 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 1)
    Q = "select * from mov_clientes where cliente= '" & Label10 & "'"
    TABLA.Open Q, conexion_BD
    sal = 0
    haber = 0
    Do While Not TABLA.EOF = True
        haber = CDbl(haber) + TABLA!pago
        TABLA.MoveNext

    Loop
    Label7 = Format(haber, "currency")
    TABLA.Close
    
    QQ = "select * from clientes_a_pagar where cliente='" & Label10 & "'"
    TABLA.Open QQ, conexion_BD
    sal = 0
    debe = 0
    Do While Not TABLA.EOF = True
        
        'If TABLA!verdadero = 0 Then
        debe = CDbl(debe) + TABLA!monto
            
        'End If
        TABLA.MoveNext

    Loop
    Label8 = Format(debe, "currency")
    resta = CDbl(Label7) - CDbl(Label8)
    Label9 = Format(resta, "currency")
    Label6 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 1)
    TABLA.Close
    Frame2.Visible = True
    
End Sub



Private Sub MSFlexGrid4_Click()
'PROVEEDORES
    Label10 = Me.MSFlexGrid4.TextMatrix(Me.MSFlexGrid4.Row, 1)
    Q = "select * from mov_proveedor where proveedor= '" & Label10 & "'"
    TABLA.Open Q, conexion_BD
    sal = 0
    debe = 0
    Do While Not TABLA.EOF = True
        debe = CDbl(debe) + TABLA!pago
        TABLA.MoveNext
    Loop
    Label7 = Format(debe, "currency")
    TABLA.Close
    
    QQ = "select * from prove_a_pagar where proveedor='" & Label10 & "'"
    TABLA.Open QQ, conexion_BD
    sal = 0
    haber = 0
    Do While Not TABLA.EOF = True
        'If TABLA!verdadero = -1 Then
        haber = CDbl(haber) + TABLA!monto
        'End If
        TABLA.MoveNext
    Loop
    TABLA.Close
    
    a = "select * from retencion where destino='" & Label10 & "'"
    TABLA.Open a, conexion_BD
    Label15 = 0
    ret = 0
    Do While Not TABLA.EOF
        ret = CDbl(ret) + TABLA!importe
        TABLA.MoveNext
    Loop
    TABLA.Close
    Label15 = ret
    Label8 = Format(haber, "currency")
    suma = CDbl(Label7) + CDbl(Label15)
    Label7 = Format(suma, "currency")
    resta = CDbl(Label7) - CDbl(Label8)
    Label9 = Format(resta, "currency")
    Label6 = Me.MSFlexGrid4.TextMatrix(Me.MSFlexGrid4.Row, 1)
   
    Frame2.Visible = True
End Sub

Private Sub MSFlexGrid5_Click()
Command9.Visible = True
''' codigo sacado de internet y copiado tal cual '''
Dim oldx, oldy, cell2text As String, strTextCheck As String

' Check or uncheck the grid checkbox
With MSFlexGrid5
    oldx = .Col
    oldy = .Row
        If .Col = 0 Then
            If .CellPicture = picChecked.Picture Then
                Set .CellPicture = picUnchecked.Picture
                .Col = .Col + 1  ' I use data that is in column #1, usually an Index or ID #
                strTextCheck = .Text
                ' When you de-select a CheckBox, we need to strip out the #
                strChecked = Replace(strChecked, strTextCheck & ",", "")
                ' Don't forget to strip off the trailing , before passing the string
                Debug.Print strChecked
                .TextMatrix(oldy, 6) = ""
            Else
                Set .CellPicture = picChecked.Picture
                '.TextMatrix(oldx, oldy) = "1"
                .Col = .Col + 1
                strTextCheck = .Text
                strChecked = strChecked & strTextCheck & ","
                Debug.Print strChecked
                .TextMatrix(oldy, 6) = "1"
            End If
        End If
    .Col = oldx
    .Row = oldy
End With
End Sub

Private Sub OptionClientes_Click()
Label14 = "CLIENTES"

MSFlexGrid1.Visible = False
MSFlexGrid4.Visible = False
MSFlexGrid5.Visible = False
MSFlexGrid2.Clear
MSFlexGrid2.Visible = True

Text3.Visible = True
Command5.Visible = False
Command6.Visible = False
Command7.Visible = True
Command4.Visible = True
Text1.Visible = False
Text2.Visible = False
Text3.Visible = True


Label6 = ""
Label7 = ""
Label8 = ""
Label9 = ""
Label10 = ""
Text4 = ""

Call ALTALISTA2

com = "select * from clientes order by nombre_cliente asc"
TABLA.Open com, conexion_BD
Frame3.Visible = True

Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!id_cliente
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!nombre_cliente

        TABLA.MoveNext
    Loop
    
    TABLA.Close
    
End Sub

Private Sub OptionPersonal_Click()
Label14 = "PERSONAL"

MSFlexGrid1.Clear
MSFlexGrid1.Visible = True
MSFlexGrid2.Visible = False
MSFlexGrid4.Visible = False
MSFlexGrid5.Visible = False
Text1.Visible = True
Text2.Visible = False
Text3.Visible = False
Command5.Visible = True
Command6.Visible = False
Command7.Visible = False

Command1.Visible = False
Command2.Visible = False
'Command4.Visible = False
Frame4.Visible = False

Label6 = ""
Label7 = ""
Label8 = ""
Label9 = ""
Label10 = ""

Call ALTALISTA
W = "select * from alta_rrhh order by nombre_rrhh asc"
TABLA.Open W, conexion_BD
Frame2.Visible = True
Frame3.Visible = True
    lin = 0
    Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!id_rrhh
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!nombre_rrhh
        TABLA.MoveNext
    Loop
    
    TABLA.Close
    
End Sub

Private Sub ALTALISTA()
'Personal

MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 2
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "ID"
MSFlexGrid1.TextMatrix(0, 1) = "APELLIDO Y NOMBRE"


MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 3000

End Sub

Private Sub ALTALISTA2()

'clients
MSFlexGrid2.FixedCols = 0
MSFlexGrid2.Cols = 2
MSFlexGrid2.FixedRows = 1
MSFlexGrid2.Rows = 2
MSFlexGrid2.Clear
MSFlexGrid2.TextMatrix(0, 0) = "ID CLIENTE"
MSFlexGrid2.TextMatrix(0, 1) = "APELLIDO Y NOMBRE"


MSFlexGrid2.ColWidth(0) = 1000
MSFlexGrid2.ColWidth(1) = 3000


End Sub
Private Sub ALTALISTA3()
'alta proveedor
MSFlexGrid4.FixedCols = 0
MSFlexGrid4.Cols = 2
MSFlexGrid4.FixedRows = 1
MSFlexGrid4.Rows = 2
MSFlexGrid4.Clear
MSFlexGrid4.TextMatrix(0, 0) = "ID PROV."
MSFlexGrid4.TextMatrix(0, 1) = "RAZON SOCIAL"

MSFlexGrid4.ColWidth(0) = 1000
MSFlexGrid4.ColWidth(1) = 3000

End Sub
Private Sub ALTALISTA4()
MSFlexGrid5.FixedCols = 0
MSFlexGrid5.Cols = 7
MSFlexGrid5.Rows = 2
MSFlexGrid5.FixedRows = 1

MSFlexGrid5.Clear
MSFlexGrid5.TextMatrix(0, 0) = ""
MSFlexGrid5.TextMatrix(0, 1) = "NOMBRE"
MSFlexGrid5.TextMatrix(0, 2) = "PAGOS"
MSFlexGrid5.TextMatrix(0, 3) = "FACTURAS"
MSFlexGrid5.TextMatrix(0, 4) = "RETENCIONES"
MSFlexGrid5.TextMatrix(0, 5) = "SALDO"


MSFlexGrid5.ColWidth(0) = 300
MSFlexGrid5.ColWidth(1) = 3000
MSFlexGrid5.ColWidth(2) = 1500
MSFlexGrid5.ColWidth(3) = 1500
MSFlexGrid5.ColWidth(4) = 1200
MSFlexGrid5.ColWidth(5) = 1800
MSFlexGrid5.ColWidth(6) = 10
End Sub
Private Sub ALTALISTA5()
MSFlexGrid5.FixedCols = 0
MSFlexGrid5.Cols = 7
MSFlexGrid5.Rows = 2
MSFlexGrid5.FixedRows = 1

MSFlexGrid5.Clear
MSFlexGrid5.TextMatrix(0, 0) = ""
MSFlexGrid5.TextMatrix(0, 1) = "NOMBRE"
MSFlexGrid5.TextMatrix(0, 2) = "ADELANTO"
MSFlexGrid5.TextMatrix(0, 3) = "PREMIO"
MSFlexGrid5.TextMatrix(0, 4) = "SUELDO"
MSFlexGrid5.TextMatrix(0, 5) = "SALDO"

MSFlexGrid5.ColWidth(0) = 300
MSFlexGrid5.ColWidth(1) = 3000
MSFlexGrid5.ColWidth(2) = 1700
MSFlexGrid5.ColWidth(3) = 1500
MSFlexGrid5.ColWidth(4) = 1700
MSFlexGrid5.ColWidth(5) = 1700
MSFlexGrid5.ColWidth(6) = 10

End Sub
Private Sub OptionProve_Click()
Label14 = "PROVEEDORES"

MSFlexGrid4.Clear
MSFlexGrid1.Visible = False
MSFlexGrid2.Visible = False
MSFlexGrid5.Visible = False
MSFlexGrid4.Visible = True
Text1.Visible = False
Text2.Visible = True
Text3.Visible = False

Command5.Visible = False
Command6.Visible = True
Command7.Visible = False
Command4.Visible = True

Label6 = ""
Label7 = ""
Label8 = ""
Label9 = ""
Label10 = ""
Text4 = ""

Call ALTALISTA3
W = "select * from proveedores order by nombre_prove asc"
TABLA.Open W, conexion_BD

Frame3.Visible = True
    lin = 0
    Do While Not TABLA.EOF
    lin = lin + 1
        MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
        MSFlexGrid4.TextMatrix(lin, 0) = TABLA!id_prove
        MSFlexGrid4.TextMatrix(lin, 1) = TABLA!nombre_prove
        TABLA.MoveNext
    Loop
    
    TABLA.Close
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Q = "select * from alta_rrhh" 'where id_rrhh= " & Val(Label10) & ""
    TABLA.Open Q, conexion_BD
    sal = 0
    MSFlexGrid1.Clear
    Call ALTALISTA
    Do While Not TABLA.EOF
    If UCase(Left(TABLA!nombre_rrhh, Len(Text1))) = UCase(Text1) Then
    
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!id_rrhh
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!nombre_rrhh
        'TABLA.MoveNext
  
    End If
    TABLA.MoveNext
    Loop
TABLA.Close

End If

End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Q = "select * from proveedores"
    TABLA.Open Q, conexion_BD
    sal = 0
    lin = 0
    MSFlexGrid4.Clear
    Call ALTALISTA3
    Do While Not TABLA.EOF
    If UCase(Left(TABLA!nombre_prove, Len(Text2))) = UCase(Text2) Then
    
        lin = lin + 1
        MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
        MSFlexGrid4.TextMatrix(lin, 0) = TABLA!id_prove
        MSFlexGrid4.TextMatrix(lin, 1) = TABLA!nombre_prove
  
    End If
    TABLA.MoveNext
    Loop
TABLA.Close

End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Q = "select * from clientes"
    TABLA.Open Q, conexion_BD
    sal = 0
    lin = 0
    MSFlexGrid2.Clear
    Call ALTALISTA2
    Do While Not TABLA.EOF
    If UCase(Left(TABLA!nombre_cliente, Len(Text3))) = UCase(Text3) Then
    
        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!id_cliente
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!nombre_cliente
        'TABLA.MoveNext

    End If
    TABLA.MoveNext
    Loop
TABLA.Close

End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With MSFlexGrid5

For a = 1 To .Rows - 1

    If .TextMatrix(a, 5) = "SALDO" Or .TextMatrix(a, 5) = "" Then
        

    Else
        If .TextMatrix(a, 5) >= Val(Text4) Then
            .Row = a
            .RemoveItem (.Row)
            a = a - 1
            .Rows = .Rows + 1
        End If
    End If

Next

End With
End If
End Sub

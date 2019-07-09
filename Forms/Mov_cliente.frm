VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mov_cliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de clientes"
   ClientHeight    =   9360
   ClientLeft      =   2085
   ClientTop       =   960
   ClientWidth     =   11775
   Icon            =   "Mov_cliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   11775
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agregar"
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
      Left            =   9840
      Picture         =   "Mov_cliente.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Mov_cliente.frx":0B14
      Left            =   8880
      List            =   "Mov_cliente.frx":0B1E
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Left            =   9840
      Picture         =   "Mov_cliente.frx":0B2A
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Listar"
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
      Left            =   9840
      Picture         =   "Mov_cliente.frx":10B4
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6480
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
      Left            =   9840
      Picture         =   "Mov_cliente.frx":163E
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cheques ingresados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   1680
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   6000
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   4680
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   3480
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   2400
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Pagos Realizados:"
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
         Left            =   1560
         TabIndex        =   68
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "C.U.I.T."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Venc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "Cheque Nº"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Interno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   120
      TabIndex        =   38
      Top             =   1800
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CheckBox Check1 
         Caption         =   "Afectar factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   42
         Top             =   1320
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   8160
         TabIndex        =   67
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61472769
         CurrentDate     =   41085
      End
      Begin VB.CommandButton Command10 
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
         Left            =   9720
         Picture         =   "Mov_cliente.frx":1BC8
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text9 
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
         Left            =   1320
         TabIndex        =   39
         Top             =   840
         Width           =   5175
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
         Left            =   5160
         TabIndex        =   41
         Text            =   "0"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Mov_cliente.frx":2152
         Left            =   1080
         List            =   "Mov_cliente.frx":2165
         TabIndex        =   40
         Top             =   1440
         Width           =   1695
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
         Left            =   4800
         TabIndex        =   44
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   7680
         TabIndex        =   45
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   1320
         TabIndex        =   46
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text8 
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
         Left            =   7560
         MaxLength       =   11
         TabIndex        =   50
         ToolTipText     =   "PRESIONAR ENTER PARA CARGAR AL LISTADO"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
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
         Left            =   9720
         Picture         =   "Mov_cliente.frx":2196
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Listar"
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
         Left            =   9720
         Picture         =   "Mov_cliente.frx":2720
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1680
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5040
         TabIndex        =   48
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61472769
         CurrentDate     =   41043
      End
      Begin VB.Label Label35 
         Caption         =   "Monto retencion:"
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
         Left            =   2880
         TabIndex        =   72
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2280
         TabIndex        =   71
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Total:"
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
         TabIndex        =   70
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label34 
         Caption         =   "Fecha pago:"
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
         Left            =   6600
         TabIndex        =   66
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
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
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pago:"
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
         TabIndex        =   63
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº interno:"
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
         TabIndex        =   62
         Top             =   2160
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de cheque:"
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
         Left            =   3000
         TabIndex        =   61
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
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
         TabIndex        =   60
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
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
         TabIndex        =   56
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de venc.:"
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
         Left            =   3120
         TabIndex        =   51
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1680
         TabIndex        =   43
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto en efectivo"
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
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "C.U.I.T:"
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
         TabIndex        =   47
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   240
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   11055
      Begin VB.TextBox Text10 
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
         Left            =   2520
         TabIndex        =   30
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   7080
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61472769
         CurrentDate     =   41082
      End
      Begin VB.CommandButton Command9 
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
         Left            =   8640
         Picture         =   "Mov_cliente.frx":2CAA
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   960
         Width           =   975
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
         Left            =   1200
         TabIndex        =   28
         Top             =   1080
         Width           =   4095
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
         TabIndex        =   25
         Top             =   360
         Width           =   1455
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
         Left            =   4200
         TabIndex        =   26
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Listar"
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
         Left            =   8640
         Picture         =   "Mov_cliente.frx":3234
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command7 
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
         Left            =   8640
         Picture         =   "Mov_cliente.frx":37BE
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7080
         TabIndex        =   27
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61472769
         CurrentDate     =   41073
      End
      Begin VB.Label Label39 
         Caption         =   "label 39     DETALLE COMPLETO"
         Height          =   495
         Left            =   4200
         TabIndex        =   75
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label Label38 
         Caption         =   "Precio Unitario: $"
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
         Left            =   240
         TabIndex        =   74
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         Caption         =   "Fecha Vto.:"
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
         Left            =   5520
         TabIndex        =   65
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Factura:"
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
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
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
         Left            =   3240
         TabIndex        =   34
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fact.:"
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
         Left            =   5520
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label36 
      Caption         =   "aca hay 2 botnes, command1 esta oculto"
      Height          =   1095
      Left            =   10920
      TabIndex        =   73
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label37 
      Caption         =   "Label37"
      Height          =   375
      Left            =   10560
      TabIndex        =   69
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label30 
      Caption         =   "L30=coonvertir"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label29 
      Caption         =   "L29= dif de pago"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "L22= monto total a pagar"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label20 
      Caption         =   "SISTEMA= VIVERO O STELLA"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label21 
      Caption         =   "Label21=n remito"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Desea cargar un pago?:"
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
      Left            =   5880
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "fecha "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMIENTOS DE CLIENTES"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "Mov_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim minimo As Long
Dim maximo As Long
Dim valido As String
Dim resta As String




Private Sub Check1_Click()

datos = Combo1.Text
sValor = "CLIENTES"
Afectar_factura.Show
End Sub


Private Sub Combo1_Click()
e = " select * from mov_clientes where cliente ='" & Combo1 & "'"
TABLA.Open e, conexion_BD
If TABLA.EOF = False Then
lin = 0
Label10 = ""
tot1 = 0
tot2 = 0
Do While Not TABLA.EOF
    lin = lin + 1
    'movimiento de clientes
    tot1 = CDbl(tot1) + TABLA!pago 'DEBE
    TABLA.MoveNext
Loop
    
TABLA.Close
    
f = "select * from clientes_a_pagar where cliente ='" & Combo1 & "'"
TABLA.Open f, conexion_BD
Do While Not TABLA.EOF
    lin = 0
    lin = lin + 1
    'clientes a pagar
    tot2 = CDbl(tot2) + TABLA!monto 'HABER
    TABLA.MoveNext
    Loop
TABLA.Close
    
resta = CDbl(tot2) - CDbl(tot1)
Label10 = Format(resta, "currency")
Label8.Visible = True
Label10.Visible = True
    
Else
    'MsgBox "El cliente no posee movimientos registrados a la fecha.", vbOKOnly, "VIVERO SAN NICOLAS SA"
    TABLA.Close
    Label8.Visible = False
    Label10.Visible = False
End If
End Sub

Private Sub Combo2_Click()
If Combo2 = "SI" Then
'CARGA UN PAGO
    Frame1.Visible = False
    Frame3.BorderStyle = 0
    Frame2.Visible = False
    Frame3.Visible = True
'TRAEMOS EL SALDO DE LA CUENTA
e = " select * from mov_clientes where cliente ='" & Combo1 & "'"
TABLA.Open e, conexion_BD
If TABLA.EOF = False Then
lin = 0
Label10 = ""
tot1 = 0
tot2 = 0
Do While Not TABLA.EOF
    lin = lin + 1
    'movimiento de clientes
    tot1 = CDbl(tot1) + TABLA!pago 'DEBE
    TABLA.MoveNext
Loop
    
TABLA.Close
    
f = "select * from clientes_a_pagar where cliente ='" & Combo1 & "'"
TABLA.Open f, conexion_BD
Do While Not TABLA.EOF
    lin = 0
    lin = lin + 1
    'clientes a pagar
    'If TABLA!verdadero = "0" Then
        tot2 = CDbl(tot2) + TABLA!monto 'HABER
    'End If
    TABLA.MoveNext
    Loop
TABLA.Close
    
resta = CDbl(tot1) - CDbl(tot2)
Label10 = Format(resta, "currency")
Label8.Visible = True
Label10.Visible = True
    
Else
    'MsgBox "El cliente no posee movimientos registrados a la fecha.", vbOKOnly, "VIVERO SAN NICOLAS SA"
    TABLA.Close
    Label8.Visible = False
    Label10.Visible = False
End If

    
    
        
Else
'CARGA UNA FACTURA

    Label8.Visible = False
    Label10.Visible = False
    
    Frame1.Visible = False
    Frame2.Visible = True
    Frame2.BorderStyle = 0
    Frame3.Visible = False
    
    Command1.Visible = False
    Command2.Visible = False
    Command8.Visible = False
    Text4 = "0"
    Text5 = "0"
    Text6 = "0"
    Text7 = "0"
    Label13 = "0"
    Label19 = "0"
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    List6.Clear
    
End If
End Sub


Private Sub Combo3_Click()

Select Case Combo3.Text
    Case "CHEQUE"
    ''' PAGO EN EFECTIVO
        Label5.Visible = False
        Text4.Visible = False
    ''' pago en cheques
        Label14.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label17.Visible = True
        Label18.Visible = True
        Label19.Visible = True
        Label7.Visible = True
        Text5.Visible = True
        Text6.Visible = True
        Text7.Visible = True
        Text8.Visible = True
        DTPicker1.Visible = True
        
        Frame1.Visible = True
    
        a = "select max(n_interno) from entracheque"
        TABLA.Open a, conexion_BD
        Label19 = TABLA.Fields(0) + 1
        TABLA.Close
        Label35.Visible = False
        Command1.Visible = False
        Command3.Visible = True
        Command2.Visible = True
        Command8.Visible = True
        Command4.Visible = False
        Command5.Visible = False
        Command10.Visible = False
                
    Case "EFECTIVO"

        Frame1.Visible = False
        List1.Clear
        List2.Clear
        List3.Clear
        List4.Clear
        List5.Clear
        List6.Clear
        Command1.Visible = False
        Command2.Visible = False
        Command8.Visible = False
        Command3.Visible = False
        
        Command4.Visible = True
        Command5.Visible = True
        Command10.Visible = True
    ''' PAGO EN EFECTIVO
        Label5 = "Monto en efectivo"
        Label5.Visible = True
        Text4.Visible = True
        Label35.Visible = False
    ''' pago en cheques
        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = False
        Label7.Visible = False
        Text5.Visible = False
        Text6.Visible = False
        Text7.Visible = False
        Text8.Visible = False
        DTPicker1.Visible = False
        
    Case "RETENCION"
              
        Frame1.Visible = False
        List1.Clear
        List2.Clear
        List3.Clear
        List4.Clear
        List5.Clear
        List6.Clear
        Command1.Visible = False
        Command2.Visible = False
        Command8.Visible = False
        Command3.Visible = False
        
        Command4.Visible = True
        Command5.Visible = True
        Command10.Visible = True
    ''' PAGO EN EFECTIVO
        Label35.Visible = True
        Label5.Visible = False
        Text4.Visible = True
    ''' pago en cheques
        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = False
        Label7.Visible = False
        Text5.Visible = False
        Text6.Visible = False
        Text7.Visible = False
        Text8.Visible = False
        DTPicker1.Visible = False
        
    Case "INSUMOS"
    ''' SE UTILIZA PARA CDO DE PAGA CON GASOIL POR EJEMPLO
        Frame1.Visible = False
        List1.Clear
        List2.Clear
        List3.Clear
        List4.Clear
        List5.Clear
        List6.Clear
        Command1.Visible = False
        Command2.Visible = False
        Command8.Visible = False
        Command3.Visible = False
        
        Command4.Visible = True
        Command5.Visible = True
        Command10.Visible = True
    ''' PAGO EN EFECTIVO
        Label5 = "Monto: "
        Label5.Visible = True
        Text4.Visible = True
        Label35.Visible = False
    ''' pago en cheques
        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = False
        Label7.Visible = False
        Text5.Visible = False
        Text6.Visible = False
        Text7.Visible = False
        Text8.Visible = False
        DTPicker1.Visible = False
    Case Else
    ''' PAGO EN EFECTIVO
        Label5 = "Monto:"
        Label35.Visible = False
        Label5.Visible = True
        Text4.Visible = True
    ''' pago en cheques
        Label14.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label17.Visible = True
        Label18.Visible = True
        Label19.Visible = True
        Label7.Visible = True
        Text5.Visible = True
        Text6.Visible = True
        Text7.Visible = True
        Text8.Visible = True
        DTPicker1.Visible = True
        Frame1.Visible = True
        a = "select max(n_interno) from entracheque"
        TABLA.Open a, conexion_BD
        Label19 = TABLA.Fields(0) + 1
        TABLA.Close
        Command3.Visible = True
        Command2.Visible = True
        Command8.Visible = True
        Command4.Visible = False
        Command5.Visible = False
        Command10.Visible = False
        
End Select
End Sub

Private Sub Combo4_Click()
Select Case Combo4.Text
    Case "PROPIOS"
        Label14.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label17.Visible = True
        Label18.Visible = True
        Label19.Visible = True
        Label21.Visible = False
        Combo5.Visible = False
        Text5.Visible = True
        Text6.Visible = True
        Text7.Visible = True
        DTPicker1.Visible = True
    
        a = "select max(n_interno) from entracheque"
        TABLA.Open a, conexion_BD
        Label19 = TABLA.Fields(0) + 1
        TABLA.Close
    Case "TERCEROS"
        Label21.Visible = True
        Combo5.Visible = True
        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = False
        Text5.Visible = False
        Text6.Visible = False
        Text7.Visible = False
        DTPicker1.Visible = False
        
    Case Else
        Label14.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label17.Visible = True
        Label18.Visible = True
        Label19.Visible = True
        Text5.Visible = True
        Text6.Visible = True
        Text7.Visible = True
        DTPicker1.Visible = True
        
        Label21.Visible = True
        Combo5.Visible = True
        
        a = "select max(n_interno) from entracheque"
        TABLA.Open a, conexion_BD
        Label19 = TABLA.Fields(0) + 1
        TABLA.Close
   
    End Select
End Sub

Private Sub Command1_Click()

For a = 0 To List1.ListCount - 1
    List1.ListIndex = a: n_interno = List1.List(a)
    List2.ListIndex = a: n_cheque = List2.List(a)
    List3.ListIndex = a: importe = List3.List(a)
    List4.ListIndex = a: banco = List4.List(a)
    List5.ListIndex = a: fecha = List5.List(a)
    List6.ListIndex = a: cuit = List6.List(a)

tot = Val(tot) + importe
Next
If Text4 = "" Then
    Text4 = 0
End If
Label22 = CDbl(tot) + CDbl(Text4) ' monto total a pagar
Label29 = CDbl(Text4) - CDbl(Label22)  'diferencia de pago

txtmon = CDbl(Label22)
Call CONVERTIR
Label30 = txtmonl

remito = "select max(n_remito) from remito_interno"
TABLA.Open remito, conexion_BD
Label21 = TABLA.Fields(0) + 1
TABLA.Close


rint = "insert into remito_interno values (" & Val(Label21) & ",'" & DTPicker4 & "','" & Combo1 & "','" & Label20 & "','" & Text9 & "','" & Label22 & "','" & usua & "')"
conexion_BD.Execute rint

If datos7 = "PARCIAL" Then ' si el pago es total se guarda desde "afectar_factura.frm"

    fra = "insert into facturas values ('" & datos1 & "','" & Combo1 & "','" & Text9 & "','0','" & DTPicker4 & "','" & datos4 & "','" & datos3 & "','PARCIAL','" & datos2 & "','" & Label21 & "')"
    conexion_BD.Execute fra
End If

Select Case Combo3
    Case "EFECTIVO"
        cli = " insert into mov_clientes values ('" & Combo1 & "','" & Text9 & "','" & Text4 & "'," & Val(0) & ",'" & DTPicker4 & "','" & Label21 & "','EFECTIVO')"
        conexion_BD.Execute cli
        CA = "insert into mov_caja values ('" & DTPicker4 & "','" & Combo1 & " " & Text9 & " " & Text2 & "','" & Text4 & "','" & 0 & "','" & Label21 & "')"
        conexion_BD.Execute CA
               
    Case "CHEQUE"
        For i = 0 To List1.ListCount - 1
            cli = " insert into mov_clientes values ('" & Combo1 & "','" & Text9 & "','" & List3.List(i) & "'," & Val(List2.List(i)) & ",'" & DTPicker4 & "','" & Label21 & "','CHEQUE')"
            conexion_BD.Execute cli
            che = "insert into entracheque values (" & Val(List1.List(i)) & "," & Val(List2.List(i)) & ",'" & List3.List(i) & "','" & Combo1 & "','" & List4.List(i) & "','" & List5.List(i) & "','" & DTPicker4 & "'," & 0 & "," & Val(Label21) & "," & Val(List6.List(i)) & ")"
            conexion_BD.Execute che

            
        Next
        
    Case "AMBOS"
        CLIefe = " insert into mov_clientes values ('" & Combo1 & "','" & Text9 & "','" & Text4 & "'," & Val(0) & ",'" & DTPicker4 & "','" & Label21 & "','EFECTIVO')"
        conexion_BD.Execute CLIefe
        For i = 0 To List1.ListCount - 1
            cli = " insert into mov_clientes values ('" & Combo1 & "','" & Text9 & "','" & List3.List(i) & "'," & Val(List2.List(i)) & ",'" & DTPicker4 & "','" & Label21 & "','CHEQUE')"
            conexion_BD.Execute cli
            che = "insert into entracheque values (" & Val(List1.List(i)) & "," & Val(List2.List(i)) & ",'" & List3.List(i) & "','" & Combo1 & "','" & List4.List(i) & "','" & List5.List(i) & "','" & DTPicker4 & "'," & 0 & "," & Val(Label21) & "," & Val(List6.List(i)) & ")"
            conexion_BD.Execute che
        Next
        CA = "insert into mov_caja values ('" & DTPicker4 & "','" & Combo1 & " " & Text9 & " " & Text2 & "','" & Text4 & "','" & 0 & "','" & Label21 & "')"
        conexion_BD.Execute CA
        

        
End Select

Call IMPRIMIR

DTPicker1.Visible = False
Command1.Visible = False
Command2.Visible = False

Command8.Visible = False
Label13 = ""
Label19 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text8 = ""
Text9 = ""
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False

List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear


End Sub

Private Sub Command10_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Label10 = ""
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear

Label8.Visible = False
Label10.Visible = False
End Sub

Private Sub Command2_Click()
datos = Combo1.Text
sValor = "CLIENTES"
Detalles_ctas.Show

End Sub

Private Sub Command3_Click()

If Text5 = Val(List2.List(i)) Then
            
    MsgBox "El cheque ya está en la lista.", vbOKOnly + vbInformation, "¡ATENCIÓN!"
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Text5.SetFocus
Else
    If DTPicker1 > minimo And DTPicker1 < maximo Then
        List1.AddItem (Label19)
        List2.AddItem (Text5)
        List3.AddItem (Text6)
        List4.AddItem (Text7)
        List5.AddItem (DTPicker1)
        List6.AddItem (Text8)
    
        m = MsgBox("Desea cargar otro cheque?", vbYesNo, "VIVERO SAN NICOLAS")
        
        If m = vbYes Then
            Label19 = Label19 + 1
            Text5 = ""
            Text6 = ""
            Text7 = ""
            Text8 = ""
            Text5.SetFocus
            Command1.Visible = False
            Command3.Visible = True
        Else
            Text5 = ""
            Text6 = ""
            Text7 = ""
            Text8 = ""
            Command1.Visible = True
            Command3.Visible = False
        End If
    Else
        MsgBox ("La fecha de vencimiento se encuentra fuera del rango")
    End If
End If

End Sub

Private Sub Command4_Click()
Label9 = DTPicker4
Label22 = CDbl(Text4) '+ CDbl(Text6)   'monto total a pagar
'Label29 = CDbl(Text3) - CDbl(Label22) ' diferencia de pago

'If Combo1 = "" Or Text1 = "" Then
 '   MsgBox "Falta cargar algunos datos"
'Else


remito = "select max(n_remito) from remito_interno"
TABLA.Open remito, conexion_BD
Label21 = TABLA.Fields(0) + 1
TABLA.Close

txtmon = CDbl(Text4)
Call CONVERTIR
Label30 = txtmonl

''' EN STELLA CAMBIA LABEL20 x label38
rint = "insert into remito_interno values (" & Val(Label21) & ",'" & DTPicker4 & "','" & Combo1 & "','" & Label20 & "','" & Text9 & "','" & Text4 & "','" & usua & "')"
conexion_BD.Execute rint

If datos7 = "PARCIAL" Then ' si el pago es total se guarda desde "afectar_factura.frm"
    fra = "insert into facturas values ('" & datos1 & "','" & Combo1 & "','" & Text9 & "','" & datos8 & "','" & DTPicker4 & "','" & datos4 & "','" & datos3 & "','PARCIAL','" & datos2 & "','" & Label21 & "')"
    conexion_BD.Execute fra
End If

Select Case Combo3.Text

Case "EFECTIVO"

        cli = " insert into mov_clientes values ('" & Combo1 & "','" & Text9 & "','" & Text4 & "'," & Val(0) & ",'" & DTPicker4 & "','" & Label21 & "','EFECTIVO')"
        conexion_BD.Execute cli
        
        CA = "insert into mov_caja values ('" & DTPicker4 & "','" & Combo1 & " " & Text9 & " " & Text2 & "','" & Text4 & "','" & 0 & "','" & Label21 & "')"
        conexion_BD.Execute CA

Case "RETENCION"
        
        cli = " insert into mov_clientes values ('" & Combo1 & "','" & Text9 & "','" & Text4 & "'," & Val(0) & ",'" & DTPicker4 & "','" & Label21 & "','EFECTIVO')"
        conexion_BD.Execute cli
        
        ret = "insert into retencion_clientes values ('" & DTPicker4 & "','" & Combo1 & "','" & Text4 & "','" & Label21 & "')"
        conexion_BD.Execute ret
        
Case "INSUMOS"
        ''' SE UTILIZA PARA LOS CDO SE PAGA POR EJEMPLO CON GASOIL
        cli = " insert into mov_clientes values ('" & Combo1 & "','" & Combo1 & " " & Text9 & "','" & Text4 & "'," & Val(0) & ",'" & DTPicker4 & "','" & Label21 & "','INSUMOS')"
        conexion_BD.Execute cli

End Select

'End If
Call IMPRIMIR

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text9 = ""
Text10 = ""
Label10 = ""
'MSFlexGrid1.Clear
'MSFlexGrid1.Visible = False
Label8.Visible = False
Label10.Visible = False
Command4.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame1.Visible = False
End Sub

Private Sub Command5_Click()

datos = Combo1.Text
sValor = "CLIENTES"
Detalles_ctas.Show

End Sub

Private Sub Command6_Click()
datos = Combo1.Text
sValor = "CLIENTES"
Detalles_ctas.Show
End Sub

Private Sub Command7_Click()
Label39 = Text1 & " " & Label38 & " " & Text10

validar = " select * from clientes_a_pagar where factura= '" & Text2 & "'"
TABLA.Open validar, conexion_BD
Do While Not TABLA.EOF
Label30 = TABLA!factura
Label37 = TABLA!cliente

TABLA.MoveNext
Loop

TABLA.Close
If Text2.Text = Val(Label30) And Combo1.Text = Label37 Then
    MsgBox " La factura ya ha sido ingresada. Por favor verificar el número. " & vbNewLine & " IMPORTANTE: si es parte de la misma factura," & vbNewLine & " se puede colocar de esta manera:" & vbNewLine & " EJ: 1234(1)"
    Text2.SetFocus
Else
    CLIaPAGAR = " insert into clientes_a_pagar values ('" & Combo1 & "','" & Text2 & "','" & Text1 & "','" & Text3 & "','" & DTPicker4 & "','" & DTPicker2 & "','" & DTPicker3 & "'," & 0 & ")"
    conexion_BD.Execute CLIaPAGAR
    
    fact = "insert into facturas values ('" & Text2 & "','" & Combo1 & "','" & Text1 & "','" & Text3 & "','" & DTPicker2 & "','" & DTPicker3 & "', '" & Text3 & "','IMPAGO','0','0')"
    conexion_BD.Execute fact
    Text2 = ""
    Text1 = ""
    Text3 = ""
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
End If
End Sub

Private Sub Command8_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Label10 = ""
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear

Label8.Visible = False
Label10.Visible = False
End Sub



Private Sub Command9_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Label10 = ""
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear

Label8.Visible = False
Label10.Visible = False
End Sub

Private Sub Form_Load()
Label9 = Date
DTPicker1 = Date
minimo = CDate(Me.DTPicker1.value) - 25
maximo = CDate(Me.DTPicker1.value) + 360
DTPicker2 = Date
DTPicker3 = Date
DTPicker4 = Date
Label20 = SISTEMA

d = "select * from clientes order by nombre_cliente"
TABLA.Open d, conexion_BD
Do While Not TABLA.EOF
    Combo1.AddItem TABLA!nombre_cliente
    TABLA.MoveNext
Loop
TABLA.Close

End Sub

Private Sub Altaflex()
'mov de clientes
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 4
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "FECHA"
MSFlexGrid1.TextMatrix(0, 1) = "DETALLE"
MSFlexGrid1.TextMatrix(0, 2) = "MONTO"
MSFlexGrid1.TextMatrix(0, 3) = "CHEQUE"



MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1000
MSFlexGrid1.ColWidth(3) = 800

End Sub
Private Sub AltaFlex2()
'clientes a pagar
MSFlexGrid2.FixedCols = 0
MSFlexGrid2.Cols = 5
MSFlexGrid2.FixedRows = 1
MSFlexGrid2.Rows = 2
MSFlexGrid2.Clear
MSFlexGrid2.TextMatrix(0, 0) = "FECHA"
MSFlexGrid2.TextMatrix(0, 1) = "FACTURA"
MSFlexGrid2.TextMatrix(0, 2) = "DETALLE"
MSFlexGrid2.TextMatrix(0, 3) = "MONTO"
MSFlexGrid2.TextMatrix(0, 4) = "VENCE"

MSFlexGrid2.ColWidth(0) = 1500
MSFlexGrid2.ColWidth(1) = 1500
MSFlexGrid2.ColWidth(2) = 1500
MSFlexGrid2.ColWidth(3) = 1000

End Sub

Private Sub Text1_GotFocus()
If Text1 = "0" Then
    Text1 = ""

End If
End Sub

Private Sub Text10_GotFocus()
If Text10 = "0" Then
    Text2 = ""
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text2_GotFocus()
If Text2 = "0" Then
    Text2 = ""

End If
End Sub


Private Sub Text3_GotFocus()
If Text3 = "0" Then
    Text3 = ""
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub


Private Sub Text4_GotFocus()
If Text4 = "0" Then
    Text4 = ""
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text5_GotFocus()
If Text5 = "0" Then
    Text5 = ""
End If
End Sub

Private Sub Text6_GotFocus()
If Text6 = "0" Then
    Text6 = ""
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text6_LostFocus()
If Combo3 = "EFEC y CHEQUE" Then
    resta = Val(Text4) - Val(Text6)
End If
End Sub



Private Sub Text7_GotFocus()
If Text7 = "0" Then
    Text7 = ""
End If
End Sub

Private Sub Text8_GotFocus()
If Text8 = "0" Then
    Text8 = ""
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    If Text5 = Val(List2.List(i)) Then
            
        MsgBox "El cheque ya está en la lista.", vbOKOnly + vbInformation, "¡ATENCIÓN!"
        Text5 = ""
        Text6 = ""
        Text7 = ""
        Text8 = ""
        Text5.SetFocus
    Else
        If DTPicker1 > minimo And DTPicker1 < maximo Then
            List1.AddItem (Label19)
            List2.AddItem (Text5)
            List3.AddItem (Text6)
            List4.AddItem (Text7)
            List5.AddItem (DTPicker1)
            List6.AddItem (Text8)
    
            m = MsgBox("Desea cargar otro cheque?", vbYesNo, "VIVERO SAN NICOLAS")
            
            If m = vbYes Then
                Label19 = Label19 + 1
                Text5 = ""
                Text6 = ""
                Text7 = ""
                Text8 = ""
                Text5.SetFocus
                Command1.Visible = True
            Else
                Text5 = ""
                Text6 = ""
                Text7 = ""
                Text8 = ""
                Command1.Visible = True
                Command3.Visible = False
            End If
        Else
            MsgBox ("La fecha de vencimiento se encuentra fuera del rango")
        End If
    End If
End If
'TABLA.Close
'End If
End Sub

Private Sub IMPRIMIR()

Printer.CurrentX = 30
Printer.CurrentY = 80
Printer.FontSize = 10
Printer.Print Tab(10); ""
Printer.Print Tab(15); "Nº de recibo: "; Label21; 'label 21= n_interno
Printer.Print Tab(110); DTPicker4
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí/mos de : "; Combo1; " ."
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos $ "; Label22;
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " (en letras "; Label30; " )"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " En concepto de: "; Text1; Text9; '" por la factura Nº: "; Text2
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Según el siguiente detalle:"
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Efectivo en $ "; Text4
Printer.Print Tab(15); ""
Printer.Print Tab(15); "INTERNO Nº "; Tab(30); "CHEQUE Nº"; Tab(47); "IMPORTE"; Tab(65); "BANCO"; Tab(87); "VENC."; Tab(101); "CUIT:"; 'Tab(115)
Printer.Print Tab(15); "------------------------------------------------------------------------------------------------------------------------------"

For a = 0 To List1.ListCount - 1
    List1.ListIndex = a: n_interno = List1.List(a)
    List2.ListIndex = a: n_cheque = List2.List(a)
    List3.ListIndex = a: importe = List3.List(a)
    List4.ListIndex = a: banco = List4.List(a)
    List5.ListIndex = a: fecha = List5.List(a)
    List6.ListIndex = a: cuit = List6.List(a)

Printer.Print Tab(15); " VT "; n_interno; Tab(30); n_cheque; Tab(47); importe; Tab(65); banco; Tab(87); fecha; Tab(99); cuit; 'Tab(115)
Next

Printer.Print Tab(15); "--------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Monto total en $ "; Label22
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(70); "Firma y aclaración"
Printer.Print Tab(15); ""
Printer.Print Tab(15); "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ";

Printer.Print Tab(10); ""
Printer.Print Tab(15); "Nº de recibo: "; Label21; 'label 21= n_interno
Printer.Print Tab(110); DTPicker4
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí/mos de : "; Combo1; " ."
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos $ "; Label22;
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " (en letras "; Label30; " )"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " En concepto de: "; Text1; Text9; '" por la factura Nº: "; Text2
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Según el siguiente detalle:"
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Efectivo en $ "; Text4
Printer.Print Tab(15); ""
Printer.Print Tab(15); "INTERNO Nº "; Tab(30); "CHEQUE Nº"; Tab(47); "IMPORTE"; Tab(65); "BANCO"; Tab(87); "VENC."; Tab(101); "CUIT:"; 'Tab(115)
Printer.Print Tab(15); "--------------------------------------------------------------------------------------------------------------------------------"

For a = 0 To List1.ListCount - 1
    List1.ListIndex = a: n_interno = List1.List(a)
    List2.ListIndex = a: n_cheque = List2.List(a)
    List3.ListIndex = a: importe = List3.List(a)
    List4.ListIndex = a: banco = List4.List(a)
    List5.ListIndex = a: fecha = List5.List(a)
    List6.ListIndex = a: cuit = List6.List(a)

Printer.Print Tab(15); " VT "; n_interno; Tab(30); n_cheque; Tab(47); importe; Tab(65); banco; Tab(87); fecha; Tab(99); cuit;  'Tab(115)"
Next

Printer.Print Tab(15); "--------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Monto total en $ "; Label22
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(70); "Firma y aclaración"
Printer.Print Tab(15); ""
Printer.Print Tab(15); "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ";

Printer.EndDoc


End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mov_rrhh 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos por Personal"
   ClientHeight    =   9300
   ClientLeft      =   2250
   ClientTop       =   960
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mov_rrhh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   10980
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      Height          =   6015
      Left            =   4560
      TabIndex        =   99
      Top             =   3120
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Continuar"
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
         Picture         =   "Mov_rrhh.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   5280
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4935
         Left            =   240
         TabIndex        =   100
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   8705
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
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   120
      TabIndex        =   80
      Top             =   3120
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox Text14 
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
         TabIndex        =   98
         Top             =   360
         Width           =   6255
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
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
         Left            =   9480
         Picture         =   "Mov_rrhh.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
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
         Left            =   7800
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
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
         Left            =   9480
         Picture         =   "Mov_rrhh.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Listado"
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
         Left            =   9480
         Picture         =   "Mov_rrhh.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command12 
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
         Left            =   8400
         Picture         =   "Mov_rrhh.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command11 
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
         Left            =   9480
         Picture         =   "Mov_rrhh.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
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
         ItemData        =   "Mov_rrhh.frx":26C6
         Left            =   4440
         List            =   "Mov_rrhh.frx":26D6
         TabIndex        =   22
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text12 
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
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
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
         TabIndex        =   27
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
      End
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
         Left            =   3840
         TabIndex        =   28
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox Combo4 
         Appearance      =   0  'Flat
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
         ItemData        =   "Mov_rrhh.frx":26FC
         Left            =   1560
         List            =   "Mov_rrhh.frx":2709
         TabIndex        =   24
         Top             =   1560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7680
         TabIndex        =   29
         Top             =   2280
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
         Format          =   22282241
         CurrentDate     =   41043
      End
      Begin VB.Label Label54 
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
         Height          =   375
         Left            =   240
         TabIndex        =   97
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto total: $"
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
         Left            =   240
         TabIndex        =   93
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Adelanto en:"
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
         Left            =   2880
         TabIndex        =   88
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques:"
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
         TabIndex        =   87
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label51 
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
         Left            =   3360
         TabIndex        =   86
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label50 
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
         Left            =   5880
         TabIndex        =   85
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label49 
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
         Left            =   240
         TabIndex        =   84
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label48 
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
         Left            =   2760
         TabIndex        =   83
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de vto:"
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
         TabIndex        =   82
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
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
         Left            =   4680
         TabIndex        =   25
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto efect."
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
         Left            =   6240
         TabIndex        =   81
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
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
      Left            =   1680
      TabIndex        =   96
      Top             =   2160
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8760
      TabIndex        =   95
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22282241
      CurrentDate     =   41082
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cheques Ingresados Propios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   62
      Top             =   6960
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ListBox List6 
         Height          =   1020
         Left            =   3840
         TabIndex        =   92
         Top             =   840
         Width           =   1095
      End
      Begin VB.ListBox List5 
         Height          =   1020
         Left            =   2880
         TabIndex        =   89
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List4 
         Height          =   1020
         Left            =   1920
         TabIndex        =   65
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   1020
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   855
      End
      Begin VB.ListBox List3 
         Height          =   1020
         Left            =   960
         TabIndex        =   63
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "Interno"
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
         Left            =   120
         TabIndex        =   70
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label37 
         Caption         =   "Cheque"
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
         Left            =   1080
         TabIndex        =   69
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "Importe"
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
         Left            =   2040
         TabIndex        =   68
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "Banco"
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
         Left            =   3000
         TabIndex        =   67
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label36 
         Caption         =   "Vencimiento"
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
         Left            =   3840
         TabIndex        =   66
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cheques Ingresados Terceros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5160
      TabIndex        =   71
      Top             =   6960
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ListBox List10 
         Height          =   1020
         Left            =   2880
         TabIndex        =   91
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List9 
         Height          =   1020
         Left            =   1920
         TabIndex        =   90
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List11 
         Height          =   1020
         Left            =   3840
         TabIndex        =   74
         Top             =   840
         Width           =   1095
      End
      Begin VB.ListBox List7 
         Height          =   1020
         Left            =   120
         TabIndex        =   73
         Top             =   840
         Width           =   855
      End
      Begin VB.ListBox List8 
         Height          =   1020
         Left            =   960
         TabIndex        =   72
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label43 
         Caption         =   "Vencimiento"
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
         Left            =   3840
         TabIndex        =   79
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label42 
         Caption         =   "Banco"
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
         Left            =   3000
         TabIndex        =   78
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label41 
         Caption         =   "Importe"
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
         Left            =   1920
         TabIndex        =   77
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "Cheque"
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
         Left            =   960
         TabIndex        =   76
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "Interno"
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
         Left            =   120
         TabIndex        =   75
         Top             =   480
         Width           =   735
      End
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
      Left            =   4920
      TabIndex        =   61
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command7 
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
      Left            =   8040
      Picture         =   "Mov_rrhh.frx":2727
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
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
      Left            =   6840
      TabIndex        =   11
      Text            =   "0"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   6840
      TabIndex        =   8
      Text            =   "0"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   4680
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hora"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fijo"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Calcular"
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
      Left            =   8040
      Picture         =   "Mov_rrhh.frx":2CB1
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
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
      Left            =   2160
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   8040
      Picture         =   "Mov_rrhh.frx":323B
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6120
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
      Left            =   8040
      Picture         =   "Mov_rrhh.frx":37C5
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
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
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
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
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   2400
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2400
      TabIndex        =   94
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label33 
      Caption         =   "L33=CONVERTIR"
      Height          =   255
      Left            =   8520
      TabIndex        =   60
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label32 
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
      Left            =   7200
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label31 
      Caption         =   "L31=id_rrhh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   59
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label30 
      Caption         =   " SISTEMA = VIVERO O STELLA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label29 
      Caption         =   "Label29=nº remito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   57
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Premio:"
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
      TabIndex        =   56
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Premio:"
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
      TabIndex        =   55
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto $:"
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
      Left            =   3600
      TabIndex        =   54
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Jornal:"
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
      TabIndex        =   53
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label27 
      Caption         =   "Label27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   52
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Label24"
      Height          =   255
      Left            =   9000
      TabIndex        =   51
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "Label23= label 18- text6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label22 
      Caption         =   "Label22 = Val(Label21) - Val(Label18)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Label20= fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   48
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
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
      Left            =   6120
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
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
      Left            =   6960
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Desea cargar un adelanto de sueldo?"
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
      TabIndex        =   46
      Top             =   2760
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Mes a liquidar:"
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
      TabIndex        =   45
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
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
      Left            =   4680
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
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
      Left            =   6240
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
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
      Left            =   3960
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio por hora:"
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
      TabIndex        =   44
      Top             =   5160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
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
      Left            =   3840
      TabIndex        =   43
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Adelanto sueldo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9120
      TabIndex        =   42
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total en $:"
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
      Left            =   4920
      TabIndex        =   41
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Hs:"
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
      Left            =   2760
      TabIndex        =   40
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Hs al 100%"
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
      Left            =   480
      TabIndex        =   39
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Hs al 50%:"
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
      Left            =   5760
      TabIndex        =   38
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hs normales:"
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
      TabIndex        =   37
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido y Nombre:"
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
      Left            =   4920
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Legajo Nº:"
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
      TabIndex        =   35
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos por Personal"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   5535
   End
End
Attribute VB_Name = "Mov_rrhh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valido As String
Private Sub VALIDAR_SUMA()
For i = 0 To List2.ListCount - 1
    If CDbl(List4.List(i)) = "0" Then
        propios = 0
    Else
    propios = CDbl(propios) + CDbl(List4.List(i))
    End If
Next

For i = 0 To List7.ListCount - 1
    If CDbl(List9.List(i)) = "0" Then
        terceros = 0
    Else
        terceros = CDbl(terceros) + CDbl(List9.List(i))
    End If
Next
suma = CDbl(propios) + CDbl(terceros)

If Text10 = "" Then
    Text10 = 0
End If
If Text4 = "" Then
    Text4 = 0
End If
If Text6 = "" Then
    Text6 = "0"
End If
suma1 = CDbl(Text6)
valido = CDbl(suma) + CDbl(suma1)

End Sub
Private Sub Borrar()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
    
Label13 = ""
Label14 = ""
Label15 = ""
Label26 = ""

    Text1.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Text5.Visible = False
    Text6.Visible = False
    Text7.Visible = False
    Text8.Visible = False
    Text9.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Label25.Visible = False
    Label26.Visible = False
    Label28.Visible = False
    Label32.Visible = False
    
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = False
    Command7.Visible = False
    Command6.Visible = False
    
    Option1.Visible = False
    Option2.Visible = False
    Option3.Visible = False
    Option4.Visible = False
    
    
    
    
    
    
    
End Sub


Private Sub Combo1_Click()
X = "select * from alta_rrhh where id_rrhh =" & Val(Combo1.Text) & ""
TABLA.Open X, conexion_BD
id = TABLA!id_rrhh
nombre = TABLA!nombre_rrhh
Label31 = id
Label32 = nombre
TABLA.Close
sal = 0
L = "select * from mov_rrhh where id_rrhh= " & Val(Label31) & ""
TABLA.Open L, conexion_BD
sal = 0
premio = 0
Do While Not TABLA.EOF = True

        If TABLA!premio = "" Then
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
Label18 = Format(saldo, "currency")
TABLA.Close

Label3.Visible = True
Label32.Visible = True
Label17.Visible = True
Option1.Visible = True
Option2.Visible = True
Command3.Visible = True

Command6.Visible = True

End Sub

Private Sub Combo3_Click()
Select Case Combo3.Text
    Case "INSUMOS"
    
        Frame2.Visible = False
        Frame1.Visible = False
        Label45.Visible = False
        Text6.Visible = False
        Combo4.Visible = False
        Label52.Visible = False
        
        Command3.Visible = True
        Command4.Visible = True
        Command6.Visible = True
        
    Case "EFECTIVO"
    
        Frame2.Visible = False
        Frame1.Visible = False
        Label45.Visible = False
        Text6.Visible = False
        Combo4.Visible = False
        Label52.Visible = False
        
        Command3.Visible = True
        Command4.Visible = True
        Command6.Visible = True
    
    Case "CHEQUE"
    
        Label45.Visible = False
        Text6.Visible = False
        Combo4.Visible = True
        Label52.Visible = True
        
        Command3.Visible = False
        Command4.Visible = False
        Command6.value = False
    
    
    Case Else '"AMBOS" SERIA CHEQUE Y EFECTIVO
        
        Label45.Visible = True
        Text6.Visible = True
        Combo4.Visible = True
        Label52.Visible = True
        
        Command3.Visible = False
        Command4.Visible = False
        Command6.value = False
        
End Select
End Sub



Private Sub Combo4_Click()
Select Case Combo4.Text
    Case "PROPIOS"
    
        Label46.Visible = True
        Label47.Visible = True
        Label48.Visible = True
        Label49.Visible = True
        Label50.Visible = True
        Label51.Visible = True
        Text10.Visible = True
        Text11.Visible = True
        Text12.Visible = True
        DTPicker1.Visible = True
        Command12.Visible = True
        
        'Combo5.Visible = False
        'Label30.Visible = False
        
        Frame2.Visible = True
        Frame1.Visible = False
        
        a = "select max(n_interno) from salecheque"
        TABLA.Open a, conexion_BD
        Label46 = TABLA.Fields(0) + 1
        TABLA.Close
    
    Case "TERCEROS"
        
        Label46.Visible = False
        Label47.Visible = False
        Label48.Visible = False
        Label49.Visible = False
        Label50.Visible = False
        Label51.Visible = False
        Text10.Visible = False
        Text11.Visible = False
        Text12.Visible = False
        DTPicker1.Visible = False
        Command12.Visible = False
        
        'Combo5.Visible = True
        'Label30.Visible = True
        
        Frame2.Visible = False
        Frame1.Visible = True
        
        C = "select * from entracheque order by fecha_vto"
        TABLA.Open C, conexion_BD
        'Combo5.Clear
        Frame4.Visible = True
        Call alta_cheque3
        Do While Not TABLA.EOF
        importe = TABLA!importe
        If importe > 0 Then
        'If TABLA!importe > "0" Or TABLA!rechazado = "0" Then
            'Combo5.AddItem TABLA!n_cheque & "     $ " & TABLA!importe & "     " & TABLA!fecha_vto
            lin = lin + 1
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
            MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
            MSFlexGrid1.TextMatrix(lin, 2) = TABLA!importe
            MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
            MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_vto
        
            If TABLA!rechazado = "-1" Then
                MSFlexGrid1.Row = lin
                MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
                lin = lin - 1
            End If
        End If
        
        'If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then

        TABLA.MoveNext
        Loop
        TABLA.Close
        
    Case Else
    
        Label46.Visible = True
        Label47.Visible = True
        Label48.Visible = True
        Label49.Visible = True
        Label50.Visible = True
        Label51.Visible = True
        Text10.Visible = True
        Text11.Visible = True
        Text12.Visible = True
        DTPicker1.Visible = True
        'Combo5.Visible = True
        'Label30.Visible = True
        Frame2.Visible = True
        Frame1.Visible = True
        Command12.Visible = True
        
        
        a = "select max(n_interno) from salecheque"
        TABLA.Open a, conexion_BD
        Label46 = TABLA.Fields(0) + 1
        TABLA.Close
        
        'Combo5.Clear
        C = "select * from entracheque order by fecha_vto"
        TABLA.Open C, conexion_BD
        'Combo5.Clear
        Frame4.Visible = True
        Call alta_cheque3
        Do While Not TABLA.EOF
        importe = TABLA!importe
        If importe > 0 Then
            'Combo5.AddItem TABLA!n_cheque & "     $ " & TABLA!importe & "     " & TABLA!fecha_vto
            lin = lin + 1
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
            MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
            MSFlexGrid1.TextMatrix(lin, 2) = TABLA!importe
            MSFlexGrid1.TextMatrix(lin, 3) = TABLA!banco
            MSFlexGrid1.TextMatrix(lin, 4) = TABLA!fecha_vto
        End If
        TABLA.MoveNext
        Loop
        TABLA.Close
        
    
End Select
End Sub



Private Sub Combo5_Click()

SEL = " select * from entracheque where n_cheque=" & Val(Combo5) & ""
TABLA.Open SEL, conexion_BD

If Val(Combo5) = Val(List9.List(i)) Then
    MsgBox "El cheque ya se encuentra en la lista"
Else
    List7.AddItem TABLA!n_interno
    List8.AddItem TABLA!n_cheque
    List9.AddItem TABLA!importe
    List10.AddItem TABLA!banco
    List11.AddItem TABLA!fecha_vto
    
TABLA.Close

Command11.Visible = True
End If

End Sub

Private Sub Command1_Click()
If Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text7 = "" Then
    MsgBox "Todos los campos son obligatorios!", vbOKOnly, "VIVERO SAN NICOLAS SA"
Else
    If Text9 = "" Then
        Text9 = "0"
    End If
    
    g = "insert into mov_rrhh values (" & Val(Label31) & ",'" & Label32 & "','" _
        & Text3 & "','" & Text4 & "','" & Text5 & "','" & Label12 & "','" _
        & Label14 & "','" & 0 & "','" & Text7 & "','" & Text2 & "','" _
        & Text7 & "','" & DTPicker2 & "','" & Text9 & "')"
    conexion_BD.Execute g
    
    'txtmon = cdbl(Text6)
    'Call CONVERTIR
    'Label33 = txtmonl
    'Call Imprimir
    'rint = "insert into remito_interno values (" & Val(Label29) & ",'" & dtpicker2 & "','" & Label30 & "','" & Label32 & "','" & Label9 & "'," & Val(Text3) & ")"
    'conexion_BD.Execute rint
 Call Borrar
End If
DTPicker2.Visible = False
'Text1.SetFocus
End Sub



Private Sub Command11_Click()
Call VALIDAR_SUMA

remito = "select max(n_remito) from remito_interno"
TABLA.Open remito, conexion_BD
Label29 = TABLA.Fields(0) + 1
TABLA.Close

If CDbl(valido) = CDbl(Text13) Then
Label23 = Label18 - Text13
If Combo3 = "AMBOS" Then ' EFECTIVO Y CHEQUE
    
    txtmon = CDbl(Text13)
    Call CONVERTIR
    Label33 = txtmonl
   ''' En STELLA cambia el label 30 x el Label30
    
    rint = "insert into remito_interno values (" & Val(Label29) & ",'" & DTPicker2 & "','" & Label30 & "','" & Label32 & "','" & Text14 & "','" & Text13 & "','" & usua & "')"
    conexion_BD.Execute rint
    MOVcaja = "insert into mov_caja values ('" & DTPicker2 & "','" & Text14 & Label32 & "','" & 0 & "','" & Text6 & "','" & Label29 & "')"
    conexion_BD.Execute MOVcaja
    'GUARDO EN MOV_RRHH EL TOTAL DEL ANTICIPO
    movRRHH = "insert into mov_rrhh values (" & Val(Label31) & ",'" & Label32 & "','" _
        & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" _
        & 0 & "','" & Text13 & "','" & Text14 & "','" & 0 & "','" _
        & 0 & "','" & DTPicker2 & "','" & 0 & "')"
    conexion_BD.Execute movRRHH
    


    Select Case Combo4.Text
        Case "TERCEROS"
        For i = 0 To List7.ListCount - 1
        CHE3 = "insert into mov_cheques values (" & Val(List7.List(i)) & "," & Val(List8.List(i)) & ",'" & DTPicker2 & "','" & List11.List(i) & "','" & Label30 & "','" & Label32 & "','" & List9.List(i) & "','" & Text14 & " " & Label32 & "')"
        conexion_BD.Execute CHE3
        'RRHH3 = " insert into mov_rrhh values ('" & Label31 & "','" & Label32 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & List9.List(i) & "','" & text14 & "','" & 0 & "','" & 0 & "','" & dtpicker2 & "')"
        'conexion_BD.Execute RRHH3
        Label21 = Val(List9.List(i)) * -1
        L = "update entracheque set importe='" & Label21 & "' where n_interno =" & List7.List(i) & ""
        conexion_BD.Execute L
       
        Next
   
   
    Case "PROPIOS"
    
    For i = 0 To List2.ListCount - 1
        SALECH = "insert into salecheque values (" & Val(List3.List(i)) & ",'" & List6.List(i) & "','" & Text14 & "','" & List4.List(i) & "','" & List5.List(i) & "','" & Label32 & "','" & DTPicker2 & "'," & Val(List2.List(i)) & "," & Val(0) & "," & Val(Label29) & ")"
        conexion_BD.Execute SALECH
   
        CHEp = "insert into mov_cheques values (" & Val(List2.List(i)) & "," & Val(List3.List(i)) & ",'" & DTPicker2 & "','" & List6.List(i) & "','" & Label30 & "','" & Label32 & "','" & List4.List(i) & "','" & Text14 & " " & Label32 & "')"
        conexion_BD.Execute CHEp

        'RRHHp = " insert into mov_rrhh values ('" & Label31 & "','" & Label32 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & List4.List(i) & "','" & text14 & "','" & 0 & "','" & 0 & "','" & dtpicker2 & "')"
        'conexion_BD.Execute RRHHp
    Next
   
    Case Else 'CHEQUES PROPIOS Y TERCEROS
        For i = 0 To List7.ListCount - 1 ' BUCLE PARA LOS TERCEROS
   
            CHE3 = "insert into mov_cheques values (" & Val(List7.List(i)) & "," & Val(List8.List(i)) & ",'" & DTPicker2 & "','" & List11.List(i) & "','" & Label30 & "','" & Label32 & "','" & List9.List(i) & "','" & Text14 & " " & Label32 & "')"
            conexion_BD.Execute CHE3
            'RRHH3 = " insert into mov_rrhh values ('" & Label31 & "','" & Label32 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & List9.List(i) & "','" & text14 & "','" & 0 & "','" & 0 & "','" & dtpicker2 & "')"
            'conexion_BD.Execute RRHH3
            Label21 = Val(List9.List(i)) * -1
            L = "update entracheque set importe='" & Label21 & "' where n_interno =" & List7.List(i) & ""
            conexion_BD.Execute L
   
        Next
        
        For i = 0 To List2.ListCount - 1 'BUCLE PARA LOS PROPIOS
            'RRHHp = " insert into mov_rrhh values ('" & Label31 & "','" & Label32 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & List4.List(i) & "','" & text14 & "','" & 0 & "','" & 0 & "','" & dtpicker2 & "')"
            'conexion_BD.Execute RRHHp
            SALECH = "insert into salecheque values (" & Val(List3.List(i)) & ",'" & List6.List(i) & "','" & Text14 & "','" & List4.List(i) & "','" & List5.List(i) & "','" & Label32 & "','" & DTPicker2 & "'," & Val(List2.List(i)) & "," & Val(0) & "," & Val(Label29) & ")"
            conexion_BD.Execute SALECH
   
            CHEp = "insert into mov_cheques values (" & Val(List2.List(i)) & "," & Val(List3.List(i)) & ",'" & DTPicker2 & "','" & List6.List(i) & "','" & Label30 & "','" & Label32 & "','" & List4.List(i) & "','" & Text14 & " " & Label32 & "')"
            conexion_BD.Execute CHEp
        Next
         
    
    End Select
Call IMPRIMIR


Else ' COMBO3 = "CHEQUE"
    
    txtmon = CDbl(Text13)
    Call CONVERTIR
    Label33 = txtmonl
   
    
    rint = "insert into remito_interno values (" & Val(Label29) & ",'" & DTPicker2 & "','" & Label30 & "','" & Label32 & "','" & Text14 & "','" & Text13 & "','" & usua & "')"
    conexion_BD.Execute rint
    movRRHH = "insert into mov_rrhh values (" & Val(Label31) & ",'" & Label32 & "','" _
    & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" _
    & 0 & "','" & Text13 & "','" & Text14 & "','" & 0 & "','" _
    & 0 & "','" & DTPicker2 & "','" & 0 & "')"
    conexion_BD.Execute movRRHH
    


    Select Case Combo4.Text
    
        Case "TERCEROS"
        
        For i = 0 To List7.ListCount - 1
        CHE3 = "insert into mov_cheques values (" & Val(List7.List(i)) & "," & Val(List8.List(i)) & ",'" & DTPicker2 & "','" & List11.List(i) & "','" & Label30 & "','" & Label32 & "','" & List9.List(i) & "','" & Text14 & " " & Label32 & "')"
        conexion_BD.Execute CHE3
        'RRHH3 = " insert into mov_rrhh values ('" & Label31 & "','" & Label32 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & List9.List(i) & "','" & text14 & "','" & 0 & "','" & 0 & "','" & dtpicker2 & "')"
        'conexion_BD.Execute RRHH3
        Label21 = Val(List9.List(i)) * -1
        L = "update entracheque set importe='" & Label21 & "' where n_interno =" & List7.List(i) & ""
        conexion_BD.Execute L
       
        Next
   
   
    Case "PROPIOS"
    For i = 0 To List2.ListCount - 1
        SALECH = "insert into salecheque values (" & Val(List3.List(i)) & ",'" & List6.List(i) & "','" & Text14 & "','" & List4.List(i) & "','" & List5.List(i) & "','" & Label32 & "','" & DTPicker2 & "'," & Val(List2.List(i)) & "," & Val(0) & "," & Val(Label29) & ")"
        conexion_BD.Execute SALECH
   
        CHEp = "insert into mov_cheques values (" & Val(List2.List(i)) & "," & Val(List3.List(i)) & ",'" & DTPicker2 & "','" & List6.List(i) & "','" & Label30 & "','" & Label32 & "','" & List4.List(i) & "','" & Text14 & " " & Label32 & "')"
        conexion_BD.Execute CHEp

        'RRHHp = " insert into mov_rrhh values ('" & Label31 & "','" & Label32 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & List4.List(i) & "','" & text14 & "','" & 0 & "','" & 0 & "','" & dtpicker2 & "')"
        'conexion_BD.Execute RRHHp
    Next
   
    Case Else 'CHEQUES PROPIOS Y TERCEROS
        For i = 0 To List7.ListCount - 1 ' BUCLE PARA LOS TERCEROS
   
            CHE3 = "insert into mov_cheques values (" & Val(List7.List(i)) & "," & Val(List8.List(i)) & ",'" & DTPicker2 & "','" & List11.List(i) & "','" & Label30 & "','" & Label32 & "','" & List9.List(i) & "','" & Text14 & " " & Label32 & "')"
            conexion_BD.Execute CHE3
            'RRHH3 = " insert into mov_rrhh values ('" & Label31 & "','" & Label32 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & List9.List(i) & "','" & text14 & "','" & 0 & "','" & 0 & "','" & dtpicker2 & "')"
            'conexion_BD.Execute RRHH3
            Label21 = Val(List9.List(i)) * -1
            L = "update entracheque set importe='" & Label21 & "' where n_interno =" & List7.List(i) & ""
            conexion_BD.Execute L
   
        Next
        
        For i = 0 To List2.ListCount - 1 'BUCLE PARA LOS PROPIOS
            'RRHHp = " insert into mov_rrhh values ('" & Label31 & "','" & Label32 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & List4.List(i) & "','" & text14 & "','" & 0 & "','" & 0 & "','" & dtpicker2 & "')"
            'conexion_BD.Execute RRHHp
            SALECH = "insert into salecheque values (" & Val(List3.List(i)) & ",'" & List6.List(i) & "','" & Text14 & "','" & List4.List(i) & "','" & List5.List(i) & "','" & Label32 & "','" & DTPicker2 & "'," & Val(List2.List(i)) & "," & Val(0) & "," & Val(Label29) & ")"
            conexion_BD.Execute SALECH
   
            CHEp = "insert into mov_cheques values (" & Val(List2.List(i)) & "," & Val(List3.List(i)) & ",'" & DTPicker2 & "','" & List6.List(i) & "','" & Label30 & "','" & Label32 & "','" & List4.List(i) & "','" & Text14 & " " & Label32 & "')"
            conexion_BD.Execute CHEp
        Next
         
    
    End Select
    
    Call IMPRIMIR

End If
Label32 = ""
Text13 = ""
Text6 = ""
Label46 = ""
Text12 = ""
Text11 = ""
Text10 = ""
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
List9.Clear
List10.Clear
List11.Clear
'Combo5.Clear
'Label30.Visible = False
'Combo5.Visible = False
Frame3.Visible = False
Frame2.Visible = False
Frame1.Visible = False
Label19.Visible = False
Label18.Visible = False
Else
MsgBox "No coinciden los montos"
End If
End Sub


Private Sub Command12_Click()
If Text12 = Val(List3.List(i)) Then
    MsgBox " El Cheque ya se encuentra en la lista"
    Text2.SetFocus
Else
List2.AddItem (Label46)
List3.AddItem (Text12)
List4.AddItem (Text11)
List5.AddItem (Text10)
List6.AddItem (DTPicker1)
m = MsgBox("Desea cargar otro cheque?", vbYesNo, "VIVERO SAN NICOLAS")
If m = vbYes Then
    Label46 = Label46 + 1
    Text10 = ""
    Text11 = ""
    Text12 = ""
    Text12.SetFocus
    Command11.Visible = True
Else
    Text10 = ""
    Text11 = ""
    Text12 = ""
    Command11.Visible = True
End If
End If
End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
L = "select * from mov_rrhh where id_rrhh= " & Val(Label31) & ""
TABLA.Open L, conexion_BD
sal = 0
premio = 0
Do While Not TABLA.EOF = True

        If TABLA!premio = "" Then
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
Label15 = Format(saldo, "currency")
TABLA.Close


End Sub

Private Sub Command3_Click()
Text1 = ""
Text6 = ""

Label9.Visible = False
Text6.Visible = False
Label18.Visible = False
Label19.Visible = False
Command3.Visible = False
Command4.Visible = False

Label3.Visible = False
Label32.Visible = False
Label17.Visible = False
Option1.Visible = False
Option2.Visible = False
Command6.Visible = False

Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub Command4_Click()
Label23 = Label18 - Text13

remito = "select max(n_remito) from remito_interno"
TABLA.Open remito, conexion_BD
Label29 = TABLA.Fields(0) + 1
TABLA.Close

Select Case Combo3.Text

Case "EFECTIVO"
If Text13 = "" Then
MsgBox "Debe completar los campos en blanco", vbOKOnly, "VIVERO SAN NICOLAS SA"
Else
    txtmon = CDbl(Text13)
    Call CONVERTIR
    Label33 = txtmonl
    
    
    rint = "insert into remito_interno values (" & Val(Label29) & ",'" & DTPicker2 & "','" & Label30 & "','" & Label32 & "','" & Text14 & "','" & Text13 & "','" & usua & "')"
    conexion_BD.Execute rint
    
    h = "insert into mov_rrhh values (" & Val(Label31) & ",'" & Label32 & "','" _
        & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" _
        & 0 & "','" & Text13 & "','" & Text14 & "','" _
        & 0 & "','" & 0 & "','" & DTPicker2 & "','" & 0 & "')"
    conexion_BD.Execute h
    j = "insert into mov_caja values ('" & DTPicker2 & "','" & Label32 & " " & Text14 & "','" & 0 & "','" & Text13 & "','" & Label29 & "')"
    conexion_BD.Execute j

End If
    Call IMPRIMIR
    
Case "INSUMOS"

If Text13 = "" Then
MsgBox "Debe completar los campos en blanco", vbOKOnly, "VIVERO SAN NICOLAS SA"
Else
    txtmon = CDbl(Text13)
    Call CONVERTIR
    Label33 = txtmonl
    
    
    rint = "insert into remito_interno values (" & Val(Label29) & ",'" & DTPicker2 & "','" & Label30 & "','" & Label32 & "','" & Text14 & "','" & Text6 & "','" & usua & "')"
    conexion_BD.Execute rint
    
    h = "insert into mov_rrhh values (" & Val(Label31) & ",'" & Label32 & "','" _
        & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" _
        & 0 & "','" & Text13 & "','" & Text14 & "','" _
        & 0 & "','" & 0 & "','" & DTPicker2 & "','" & 0 & "')"
    conexion_BD.Execute h

End If
    Call IMPRIMIR
    
End Select

'Call Borrar
Text6 = ""
Text13 = ""
Label18 = ""

Label19.Visible = False
Label18.Visible = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False

End Sub

Private Sub Command5_Click()
If Label14 = "" Then
    MsgBox "Faltan completar algunos campos! Por favor intentelo nuevamente!", vbOKOnly, "VIVERO SAN NICOLAS"
Else
    Label15 = Val(Label14) + Val(Label15) ' saldo
    Command1.Visible = True
End If
End Sub

Private Sub Command6_Click()
datos = Val(Label31)
datos1 = Label32
mov_jornales.Show
'Unload Me
End Sub

Private Sub Command7_Click()
If Text1 = "" Then
    Text1 = "0"
End If

If Text8 = "" Then
    Text8 = "0"
End If
'Label21 = Text1 + Text8
'saldo = CDbl(Label21) + CDbl(Label15)
'Label27 = saldo
g = "insert into mov_rrhh values (" & Val(Label31) & ",'" & Label32 & "','" _
        & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" _
        & Text1 & "','" & 0 & "','" & Text7 & "','" & 0 & "','" _
        & Text7 & "','" & DTPicker2 & "','" & Text8 & "')"
    conexion_BD.Execute g
    
    'txtmon = cdbl(ss)
    'Call CONVERTIR
    'Label33 = txtmonl
    'Call Imprimir
    'rint = "insert into remito_interno values (" & Val(Label29) & ",'" & dtpicker2 & "','" & Label30 & "','" & Label32 & "','" & Label9 & "'," & Val(Text6) & ")"
    'conexion_BD.Execute rint
Text1 = ""
Text8 = ""
Text9 = ""
Label20 = ""
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
DTPicker2.Visible = False

Call Borrar
Mov_rrhh.Refresh
End Sub

Private Sub Command8_Click()
Command11.Visible = True
Frame4.Visible = False


End Sub

Private Sub Form_Load()
Label20 = Now
Label24 = Date
DTPicker2 = Date
DTPicker1 = Date
Label30 = SISTEMA


a = "select * from alta_rrhh order by nombre_rrhh"
TABLA.Open a, conexion_BD
Do While Not TABLA.EOF
    Combo1.AddItem TABLA!id_rrhh & " " & TABLA!nombre_rrhh
    TABLA.MoveNext
Loop
TABLA.Close

End Sub
Private Sub alta_cheque3()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 5
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = " Nº INTERNO"
MSFlexGrid1.TextMatrix(0, 1) = "Nº CHEQUE"
MSFlexGrid1.TextMatrix(0, 2) = "IMPORTE"
MSFlexGrid1.TextMatrix(0, 3) = "BANCO"
MSFlexGrid1.TextMatrix(0, 4) = "VENCIMIENTO"

MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.ColWidth(4) = 1500

End Sub



Private Sub MSFlexGrid1_Click()
interno = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
cheque = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
importe = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
banco = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
vto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4)

List7.AddItem interno
List8.AddItem cheque
List9.AddItem importe
List10.AddItem banco
List11.AddItem vto

r = MsgBox("Desea cargar otro cheque?", vbYesNo)
If r = vbYes Then
    MSFlexGrid1.RemoveItem (MSFlexGrid1.RowSel)
    'If List10.List(i) = interno Then
    'MsgBox "El cheque ya ha sido ingresado"
    'End If
    Frame4.Visible = True
Else
    Command11.Visible = True
    Frame4.Visible = False
    
End If
End Sub

Private Sub Option1_Click()
If Option1 = True Then
    
    DTPicker2.Visible = True
        
    Frame3.Visible = True
    Frame3.BorderStyle = 0

    Label18.Visible = True
    Label19.Visible = True
    Command3.Visible = True
    Command6.Visible = True
    
    Text1.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Text5.Visible = False
    Text7.Visible = False
    Text8.Visible = False
    Text9.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Visible = False

    Label14.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label25.Visible = False
    Label26.Visible = False
    Label28.Visible = False
    Option3.Visible = False
    Option4.Visible = False
    
    Command1.Visible = False
    Command2.Visible = False
    Command5.Visible = False
    
L = "select * from mov_rrhh where id_rrhh= " & Val(Label31) & ""
TABLA.Open L, conexion_BD
sal = 0
premio = 0
Do While Not TABLA.EOF = True

        If TABLA!premio = "" Then
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
Label18 = Format(saldo, "currency")
TABLA.Close
End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then

'   // JORNALES COMPLETOS //
'  --------------------------
    DTPicker2.Visible = True
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Text1.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Text5.Visible = False
    Text7.Visible = False
    Text8.Visible = False
    Text9.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    Label14.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label25.Visible = False
    Label26.Visible = False
    Label28.Visible = False
    Command1.Visible = False
    Command5.Visible = False
    Command2.Visible = False
    Command7.Visible = False
    
    '// TIPO DE JORNAL //
    '-----------------------
    Label13.Visible = True
    Option3.Visible = True
    Option4.Visible = True
    
    '// ADELANTOS DE SUELDOS //
    '--------------------------

    Text6.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Command6.Visible = False
       
L = "select * from mov_rrhh where id_rrhh= " & Val(Label31) & ""
TABLA.Open L, conexion_BD
sal = 0
premio = 0
Do While Not TABLA.EOF = True

        If TABLA!premio = "" Then
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
Label15 = Format(saldo, "currency")
TABLA.Close
End If
End Sub


Private Sub Option3_Click()
If Option3 = True Then

'   // JORNALES FIJOS //
'  ---------------------

    Text1.Visible = True
    Text7.Visible = True
    Text8.Visible = True
    Label10.Visible = True
    Label13.Visible = True
    Label15.Visible = True
    Label16.Visible = True
    Label25.Visible = True
    Label26.Visible = True
    Option3.Visible = True
    Option4.Visible = True
    Command7.Visible = True
    Command2.Visible = True
       
    
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Text5.Visible = False
    Text9.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    Label14.Visible = False
    Label28.Visible = False
    Command5.Visible = False
    'Command2.Visible = False
    Label9.Visible = False
    Text6.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Command3.Visible = False
    Command4.Visible = False

End If
End Sub

Private Sub Option4_Click()
If Option4 = True Then
    Text1.Visible = False
    Text2.Visible = True
    Text3.Visible = True
    Text4.Visible = True
    Text5.Visible = True
    Text7.Visible = True
    Text8.Visible = False
    Text9.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    Label13.Visible = True
    Label12.Visible = True
    Label14.Visible = True
    Label15.Visible = True
    Label16.Visible = True
    Label25.Visible = False
    Label26.Visible = False
    Label28.Visible = True
    Command5.Visible = True
    Command2.Visible = True
    Option3.Visible = True
    Option4.Visible = True
    Command7.Visible = False
        

    Text6.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Command3.Visible = False
    Command4.Visible = False

End If
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text2_Change()
horas = Val(Text3) + Val(Text4) + Val(Text5)
Label12 = horas

tot = Val(Text2) * Val(Text3) 'valor hora normal

Y = Val(Text2) / 2
yx = Y + Val(Text2)
tot2 = Val(Text4) * yx 'valor hora al 50%

yy = Val(Text2) * 2
tot3 = Val(Text5) * yy ' valor hora al 100%

Label14 = tot + tot2 + tot3 ' suma de precio de horas

'Label15 = Val(Label14) - Val(Label15) ' saldo
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text3_Change()
horas = Val(Text3) + Val(Text4) + Val(Text5)
Label12 = horas

tot = Val(Text2) * Val(Text3) 'valor hora normal

Y = Val(Text2) / 2
yx = Y + Val(Text2)
tot2 = Val(Text4) * yx 'valor hora al 50%

yy = Val(Text2) * 2
tot3 = Val(Text5) * yy ' valor hora al 100%

Label14 = tot + tot2 + tot3 ' suma de precio de horas

'Label15 = Val(Label14) - Val(Label15) ' saldo
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text4_Change()
horas = Val(Text3) + Val(Text4) + Val(Text5)
Label12 = horas

tot = Val(Text2) * Val(Text3) 'valor hora normal

Y = Val(Text2) / 2
yx = Y + Val(Text2)
tot2 = Val(Text4) * yx 'valor hora al 50%

yy = Val(Text2) * 2
tot3 = Val(Text5) * yy ' valor hora al 100%

Label14 = tot + tot2 + tot3 ' suma de precio de horas

'Label15 = Val(Label14) - Val(Label15) ' saldo
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text5_Change()
horas = Val(Text3) + Val(Text4) + Val(Text5)
Label12 = horas

tot = Val(Text2) * Val(Text3) 'valor hora normal

Y = Val(Text2) / 2
yx = Y + Val(Text2)
tot2 = Val(Text4) * yx 'valor hora al 50%

yy = Val(Text2) * 2
tot3 = Val(Text5) * yy ' valor hora al 100%

Label14 = tot + tot2 + tot3 ' suma de precio de horas

'Label15 = Val(Label14) - Val(Label15) ' saldo

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text6_Change()
ad = Val(Text6) * 1
Label21 = ad
Label22 = CDbl(Label18) - CDbl(Label21)
If Text6 = "" Then
    Text6 = "0"
End If
Label23 = CDbl(Label18) - CDbl(Text6)

End Sub


Private Sub IMPRIMIR()


Printer.CurrentX = 30
Printer.CurrentY = 80
Printer.FontSize = 10
Printer.PaperSize = 9

Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); "Nº de recibo: "; Label29;
Printer.Print Tab(110); DTPicker2
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí/mos de: "; SISTEMA
'Printer.Print Tab(15); " Recibí/mos de: STELLA DAVIRE"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos $ "; Text13
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " ( "; Label33; " )"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " En concepto de: "; Text14;
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Según el siguiente detalle:"
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Efectivo en $ "; Text13 '; Label21
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Interno Nº "; Tab(35); "Cheque Nº"; Tab(55); "Importe"; Tab(75); "Banco"; Tab(98); "Vencimiento"; ' Tab(95); '"CUIT:"; Tab(100)
Printer.Print Tab(15); "--------------------------------------------------------------------------------------------------------------------------------"
tot = 0
tot2 = 0

For a = 0 To List4.ListCount - 1
    ' LISTADO DE CHEQUES PROPIOS
    List2.ListIndex = a: n_interno = List2.List(a)
    List3.ListIndex = a: n_cheque = List3.List(a)
    List4.ListIndex = a: importe = List4.List(a)
    List5.ListIndex = a: banco = List5.List(a)
    List6.ListIndex = a: fecha = List6.List(a)

Printer.Print Tab(15); " CH "; n_interno; Tab(35); n_cheque; Tab(55); importe; Tab(75); banco; Tab(98); fecha; ' Tab(95); 'CUIT; Tab(100)
Next

For b = 0 To List10.ListCount - 1
    'LISTADO DE CHEQUES DE TERCEROS
    
    List7.ListIndex = b: n_internoB = List7.List(b)
    List8.ListIndex = b: n_chequeB = List8.List(b)
    List9.ListIndex = b: importeB = List9.List(b)
    List10.ListIndex = b: bancoB = List10.List(b)
    List11.ListIndex = b: fechaB = List11.List(b)
Printer.Print Tab(15); " VT "; n_internoB; Tab(35); n_chequeB; Tab(55); importeB; Tab(75); bancoB; Tab(98); fechaB; ' Tab(85); 'CUIT; Tab(100)
Next

Printer.Print Tab(15); "--------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Monto total en $ "; Text13 '; Label21
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(70); Label32
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -";

'''''''''''''''''''
'' IMPRIME COPIA ''
'''''''''''''''''''

Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); "Nº de recibo: "; Label29;
Printer.Print Tab(110); DTPicker2
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí/mos de: "; SISTEMA
'Printer.Print Tab(15); " Recibí/mos de: STELLA DAVIRE"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos $ "; Text13 ' Label21
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " ( "; Label33; " )"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " En concepto de: "; Text14; '" por la factura Nº: "; Text1
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Según el siguiente detalle:"
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Efectivo en $ "; Text13 '; Label21
Printer.Print Tab(15); ""
Printer.Print Tab(15); "INTERNO Nº "; Tab(35); "CHEQUE Nº"; Tab(55); "IMPORTE"; Tab(75); "BANCO"; Tab(98); "VENCIMIENTO"; ' Tab(95); '"CUIT:"; Tab(100)
Printer.Print Tab(15); "--------------------------------------------------------------------------------------------------------------------------------"
tot = 0
tot2 = 0

For a = 0 To List4.ListCount - 1
    ' LISTADO DE CHEQUES PROPIOS
    List2.ListIndex = a: n_interno = List2.List(a)
    List3.ListIndex = a: n_cheque = List3.List(a)
    List4.ListIndex = a: importe = List4.List(a)
    List5.ListIndex = a: banco = List5.List(a)
    List6.ListIndex = a: fecha = List6.List(a)

Printer.Print Tab(15); " CH "; n_interno; Tab(35); n_cheque; Tab(55); importe; Tab(75); banco; Tab(98); fecha; ' Tab(95); 'CUIT; Tab(100)
Next

For b = 0 To List10.ListCount - 1
    'LISTADO DE CHEQUES DE TERCEROS
    
    List7.ListIndex = b: n_internoB = List7.List(b)
    List8.ListIndex = b: n_chequeB = List8.List(b)
    List9.ListIndex = b: importeB = List9.List(b)
    List10.ListIndex = b: bancoB = List10.List(b)
    List11.ListIndex = b: fechaB = List11.List(b)
Printer.Print Tab(15); " VT "; n_internoB; Tab(35); n_chequeB; Tab(55); importeB; Tab(75); bancoB; Tab(98); fechaB; ' Tab(85); 'CUIT; Tab(100)
Next

Printer.Print Tab(15); "--------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Monto total en $ "; Text13 '; Label21
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(70); Label32
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -";

Printer.EndDoc
End Sub



Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub


Private Sub Text8_GotFocus()
If Text8 = "0" Then
    Text8 = ""
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

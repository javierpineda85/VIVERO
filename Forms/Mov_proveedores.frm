VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mov_proveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moviemientos de Proveedores"
   ClientHeight    =   10290
   ClientLeft      =   1560
   ClientTop       =   435
   ClientWidth     =   11985
   Icon            =   "Mov_proveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   11985
   Begin VB.Frame Frame4 
      Height          =   6015
      Left            =   4920
      TabIndex        =   82
      Top             =   1560
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton Command4 
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
         Left            =   3000
         Picture         =   "Mov_proveedores.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   5160
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4815
         Left            =   240
         TabIndex        =   83
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   8493
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
   Begin VB.Frame Frame5 
      Height          =   4575
      Left            =   240
      TabIndex        =   65
      Top             =   2040
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox Text11 
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
         Left            =   1080
         TabIndex        =   13
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Afectar Factura"
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
         Left            =   9720
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   8280
         TabIndex        =   12
         Top             =   1080
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
         Format          =   22216705
         CurrentDate     =   41085
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
         Left            =   3240
         TabIndex        =   15
         Top             =   1800
         Width           =   975
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
         Left            =   1320
         TabIndex        =   11
         Top             =   1080
         Width           =   5175
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
         Left            =   10080
         Picture         =   "Mov_proveedores.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3000
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
         Left            =   10080
         Picture         =   "Mov_proveedores.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
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
         Left            =   10080
         Picture         =   "Mov_proveedores.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   10200
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   735
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
         ItemData        =   "Mov_proveedores.frx":1BB2
         Left            =   1560
         List            =   "Mov_proveedores.frx":1BBF
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   3720
         TabIndex        =   22
         Top             =   3240
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   1320
         TabIndex        =   21
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   7800
         TabIndex        =   20
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
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
         ItemData        =   "Mov_proveedores.frx":1BDD
         Left            =   5040
         List            =   "Mov_proveedores.frx":1BF0
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
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
         Left            =   8280
         TabIndex        =   17
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
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
         Left            =   10080
         Picture         =   "Mov_proveedores.frx":1C1F
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command6 
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
         Left            =   9120
         Picture         =   "Mov_proveedores.frx":21A9
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3720
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8520
         TabIndex        =   23
         Top             =   3240
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
         Format          =   22216705
         CurrentDate     =   41043
      End
      Begin VB.Label Label28 
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
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label27 
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
         Height          =   375
         Left            =   1200
         TabIndex        =   80
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label41 
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
         Left            =   6720
         TabIndex        =   79
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label39 
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
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Left            =   6720
         TabIndex        =   76
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label19 
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
         Left            =   4800
         TabIndex        =   75
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de vencimiento:"
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
         TabIndex        =   74
         Top             =   3240
         Visible         =   0   'False
         Width           =   3015
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
         Left            =   2760
         TabIndex        =   73
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
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
         Left            =   240
         TabIndex        =   72
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   6000
         TabIndex        =   71
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   3480
         TabIndex        =   70
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Reten.:"
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
         Left            =   2280
         TabIndex        =   69
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
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
         TabIndex        =   68
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   4320
         TabIndex        =   67
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label23 
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
         Left            =   240
         TabIndex        =   66
         Top             =   1800
         Width           =   855
      End
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
      ItemData        =   "Mov_proveedores.frx":2733
      Left            =   8640
      List            =   "Mov_proveedores.frx":273D
      TabIndex        =   2
      Top             =   1680
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
      ItemData        =   "Mov_proveedores.frx":2749
      Left            =   1680
      List            =   "Mov_proveedores.frx":274B
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Frame Frame1 
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
      Height          =   3255
      Left            =   240
      TabIndex        =   37
      Top             =   6720
      Visible         =   0   'False
      Width           =   5655
      Begin VB.ListBox List8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   4320
         TabIndex        =   42
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox List7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   3240
         TabIndex        =   41
         Top             =   840
         Width           =   1095
      End
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
         Height          =   2220
         Left            =   2160
         TabIndex        =   40
         Top             =   840
         Width           =   1095
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
         Height          =   2220
         Left            =   960
         TabIndex        =   39
         Top             =   840
         Width           =   1215
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
         Height          =   2220
         Left            =   0
         TabIndex        =   38
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label32 
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
         Left            =   240
         TabIndex        =   47
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label31 
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
         TabIndex        =   46
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label30 
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
         Left            =   2280
         TabIndex        =   45
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label29 
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
         Left            =   3360
         TabIndex        =   44
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label25 
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
         Left            =   4320
         TabIndex        =   43
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cheques Ingresados de Terceros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5880
      TabIndex        =   48
      Top             =   6720
      Visible         =   0   'False
      Width           =   5655
      Begin VB.ListBox List10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   0
         TabIndex        =   53
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   960
         TabIndex        =   52
         Top             =   840
         Width           =   1215
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
         Height          =   2220
         Left            =   2160
         TabIndex        =   51
         Top             =   840
         Width           =   1095
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
         Height          =   2220
         Left            =   3240
         TabIndex        =   50
         Top             =   840
         Width           =   1095
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
         Height          =   2220
         Left            =   4320
         TabIndex        =   49
         Top             =   840
         Width           =   1215
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
         Left            =   4320
         TabIndex        =   58
         Top             =   480
         Width           =   1095
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
         Left            =   3360
         TabIndex        =   57
         Top             =   480
         Width           =   735
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
         Left            =   2280
         TabIndex        =   56
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label33 
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
         Left            =   1200
         TabIndex        =   55
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label24 
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
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   120
      TabIndex        =   60
      Top             =   2040
      Visible         =   0   'False
      Width           =   11535
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   7800
         TabIndex        =   7
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
         Format          =   22216705
         CurrentDate     =   41082
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
         Left            =   4560
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
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
         Left            =   9480
         Picture         =   "Mov_proveedores.frx":274D
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command8 
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
         Picture         =   "Mov_proveedores.frx":2CD7
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command7 
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
         Picture         =   "Mov_proveedores.frx":3261
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7800
         TabIndex        =   5
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
         Format          =   22216705
         CurrentDate     =   41073
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
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
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
         Left            =   1320
         TabIndex        =   6
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label40 
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
         Left            =   6120
         TabIndex        =   78
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label38 
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
         Height          =   255
         Left            =   6120
         TabIndex        =   64
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Factura Nº:"
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
         TabIndex        =   63
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         TabIndex        =   62
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
         Left            =   3600
         TabIndex        =   61
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label37 
      Caption         =   "Label37"
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "SISTEMA= VIVERO O STELLA"
      Height          =   375
      Left            =   9240
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "L20= Monto total"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "L10=remito"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   5640
      TabIndex        =   33
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label26 
      Caption         =   "L26= importe *-1"
      Height          =   255
      Left            =   9360
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDORES"
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
      Left            =   5760
      TabIndex        =   31
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   9480
      TabIndex        =   30
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
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
      TabIndex        =   29
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMIENTOS DE "
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
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "Mov_proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim minimo As Long
Dim maximo As Long
Dim valido As String
Dim opcion As String
Dim i As Integer
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


Private Sub Check1_Click()
datos = Combo1.Text
sValor = "PROVEEDORES"
Afectar_factura.Show
End Sub


Private Sub Combo1_Click()
    j = "select * from mov_proveedor where proveedor= '" & Combo1 & "'"
    TABLA.Open j, conexion_BD

    Do While Not TABLA.EOF
        tot2 = CDbl(tot2) + TABLA!pago

        TABLA.MoveNext
    Loop
    TABLA.Close

    K = "select * from prove_a_pagar where proveedor= '" & Combo1 & "'"
    TABLA.Open K, conexion_BD

    Do While Not TABLA.EOF
        tot1 = CDbl(tot1) + TABLA!monto
        TABLA.MoveNext
    Loop
    TABLA.Close
    
    a = "select * from retencion where destino='" & Combo1 & "'"
    TABLA.Open a, conexion_BD
    tot3 = 0
    Do While Not TABLA.EOF

        tot3 = CDbl(tot3) + TABLA!importe
        TABLA.MoveNext
    Loop
    TABLA.Close
    suma = tot2 + tot3
    resta = suma - tot1
    Label27 = Format(resta, "currency")
    Label27.Visible = True
    Label28.Visible = True
    Command3.Visible = True
End Sub

Private Sub Combo2_Click()
If Combo2 = "SI" Then
    Frame3.Visible = False
    Frame5.Visible = True
    Frame5.BorderStyle = 0

    Command1.Visible = True
    Command3.Visible = True
    Frame1.Visible = False
    
    Frame2.Visible = False

    
    
Else
    Frame5.Visible = False
    Frame3.Visible = True
    Frame3.BorderStyle = 0
    Call NO
    Command1.Visible = True
    Command3.Visible = True
    Label3.Visible = True
    Text1.Visible = True
    Label11.Visible = True
'    Text3(0).Visible = True
    Label8.Visible = True
    
    'insert into prove a pagar
    
    
End If
End Sub

Private Sub Combo5_Click()
SEL = " select * from entracheque where n_cheque=" & Val(Combo5) & ""
TABLA.Open SEL, conexion_BD

If Val(Combo5) = Val(List9.List(i)) Then
    MsgBox "El cheque ya se encuentra en la lista"
    
Else
    List10.AddItem TABLA!n_interno
    List9.AddItem TABLA!n_cheque
    List3.AddItem TABLA!importe
    List2.AddItem TABLA!banco
    List1.AddItem TABLA!fecha_vto
    
'TABLA.Close
Command5.Visible = True
End If

TABLA.Close
End Sub

Private Sub Command1_Click()
'If Check1 = True Then
'    Afectar_factura.Show
 '   Else
    Label20 = Text11
If Text10 = "" Then
    Text10 = 0
End If
    
remito = "select max(n_remito) from remito_interno"
TABLA.Open remito, conexion_BD
Label10 = TABLA.Fields(0) + 1
TABLA.Close
    
txtmon = CDbl(Label20)
Call CONVERTIR
Label37 = txtmonl

''' en STELLA cambia el label22 x el label21
rint = "insert into remito_interno values (" & Val(Label10) & ",'" & DTPicker4 & "','" & Label22 & "','" & Combo1 & "','" & Text8 & "','" & Label20 & "','" & usua & "')"
conexion_BD.Execute rint

If datos7 = "PARCIAL" Then ' si el pago es total se guarda desde "afectar_factura.frm"

    fac = "insert into facturas values('" & datos1 & "','" & Combo1 & "','" & Text8 & "','" & datos8 & "','" & DTPicker4 & "','" & datos4 & "','" & datos3 & "','PARCIAL','" & datos2 & "','" & Label10 & "')"
    conexion_BD.Execute fac
End If

Select Case Combo3.Text
Case "EFECTIVO"
' GUARDA LOS DATOS EN EFECTIVO SOLAMENTE

    
    a = "insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & Text11 & "','" & 0 & "','" & DTPicker4 & "','" & Label10 & "','EFECTIVO')"
    conexion_BD.Execute a
    
    n = "insert into mov_caja values ('" & DTPicker4 & "','" & Combo1 & " " & Text8 & "','" & 0 & "','" & Text11 & "','" & Label10 & "')"
    conexion_BD.Execute n
       
    ret = "insert into retencion values ('" & DTPicker4 & "','" & Combo1 & "','" & Text10 & "'," & Val(Label10) & ")"
    conexion_BD.Execute ret

Case "TARJETA"
    
    a = "insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & Text11 & "','" & 0 & "','" & DTPicker4 & "','" & Label10 & "','TARJETA')"
    conexion_BD.Execute a
    
    tar = "insert into tarjeta values ('" & DTPicker4 & "','" & Combo1 & "','" & 0 & "','" & Text11 & "','" & Label10 & "')"
    conexion_BD.Execute tar
       
    ret = "insert into retencion values ('" & DTPicker4 & "','" & Combo1 & "','" & Text10 & "'," & Val(Label10) & ")"
    conexion_BD.Execute ret

Case "INSUMOS"
    ''' SE UTILIZA PARA LOS CDO SE PAGA POR EJEMPLO CON GASOIL
    a = "insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & Text11 & "','" & 0 & "','" & DTPicker4 & "','" & Label10 & "','INSUMOS')"
    conexion_BD.Execute a
    
    ret = "insert into retencion values ('" & DTPicker4 & "','" & Combo1 & "','" & Text10 & "'," & Val(Label10) & ")"
    conexion_BD.Execute ret
    
End Select

Call IMPRIMIR
Mov_proveedores.Refresh
'Call NO
Label3.Visible = True
Text1.Visible = True
Label11.Visible = True
Text11 = ""
Text1 = ""
Text2 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""

Label14 = ""
Label15 = ""
Combo5.Clear
Command2.Visible = True

End Sub

Private Sub Command2_Click()
datos = Combo1.Text
sValor = "PROVEEDORES"
Detalles_ctas.Show

End Sub



Private Sub Command3_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
List9.Clear
List10.Clear
Text1 = ""
Text2 = ""
Text9 = ""
Text4 = ""
Text5 = ""
Combo5.Clear

'MSFlexGrid2.Clear
'Frame4.Visible = False
'Command5.Visible = False
'Frame1.Visible = False
'Frame2.Visible = False
End Sub


Private Sub Command4_Click()
Command5.Visible = True
Frame4.Visible = False
End Sub

Private Sub Command5_Click()
'If Check1 = True Then
'    Afectar_factura.Show
'Else

monto = CDbl(Text11)
Call VALIDAR_SUMA

If CDbl(valido) = CDbl(monto) Then

If Text10 = "" Then
    Text10 = 0
End If

remito = "select max(n_remito) from remito_interno"
TABLA.Open remito, conexion_BD
Label10 = TABLA.Fields(0) + 1
TABLA.Close

Label9 = DTPicker4
Label20 = Text11
If Combo3 = "AMBOS" Then
    'txtmon = cdbl(Label20)
    'Call CONVERTIR
    'Label37 = txtmonl
    
    
    rint = "insert into remito_interno values (" & Val(Label10) & ",'" & DTPicker4 & "','" & Label22 & "','" & Combo1 & "','" & Text8 & "','" & Label20 & "','" & usua & "')"
    conexion_BD.Execute rint
    
    n = "insert into mov_caja values ('" & DTPicker4 & " ','" & Text8 & " " & Combo1 & "'," & 0 & ",'" & Text4 & "','" & Label10 & "')"
    conexion_BD.Execute n
    
    If datos7 = "PARCIAL" Then ' si el pago es total se guarda desde "afectar_factura.frm"
    
        fac = "insert into facturas values('" & datos1 & "','" & Combo1 & "','" & Text8 & "','" & datos8 & "','" & DTPicker4 & "','" & datos4 & "','" & datos3 & "','PARCIAL','" & datos2 & "','" & Label10 & "')"
        conexion_BD.Execute fac
    End If
    
    ' GUARDA EN MOV PROVE LOS DATOS DEL EFECTIVO
    prove = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & Text4 & "'," & Val(0) & ",'" & DTPicker4 & "','" & Label10 & "','EFECTIVO')"
    conexion_BD.Execute prove

    ret = "insert into retencion values ('" & DTPicker4 & "','" & Combo1 & "','" & Text10 & "'," & Val(Label10) & ")"
    conexion_BD.Execute ret

    Select Case Combo4.Text
        Case "TERCEROS"
        For i = 0 To List10.ListCount - 1
        CHE3 = "insert into mov_cheques values (" & Val(List10.List(i)) & "," & Val(List9.List(i)) & ",'" & DTPicker4 & "','" & List1.List(i) & "','" & 0 & "','" & Combo1 & "','" & List3.List(i) & "','" & Label6 & " " & Combo1 & " " & Label3 & " " & Val(Text1) & "')"
        conexion_BD.Execute CHE3
        
        PROV3 = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & List3.List(i) & "'," & Val(List9.List(i)) & ",'" & DTPicker4 & "','" & Label10 & "','CHEQUE VT')"
        conexion_BD.Execute PROV3
        Label26 = Val(List3.List(i)) * -1
        L = "update entracheque set importe='" & Label26 & "' where n_interno =" & List10.List(i) & ""
        conexion_BD.Execute L
        
        Next
    
        
    Case "PROPIOS"
        For i = 0 To List4.ListCount - 1
        SALECH = "insert into salecheque values (" & Val(List5.List(i)) & ",'" & List8.List(i) & "','" & Text8 & "','" & List6.List(i) & "','" & List7.List(i) & "','" & Combo1 & "','" & DTPicker4 & "'," & Val(List4.List(i)) & "," & Val(0) & "," & Val(Label10) & ")"
        conexion_BD.Execute SALECH
        
        CHEp = "insert into mov_cheques values (" & Val(List4.List(i)) & "," & Val(List5.List(i)) & ",'" & DTPicker4 & "','" & List8.List(i) & "','" & 0 & "','" & Combo1 & "','" & List6.List(i) & "','" & Label6 & " " & Combo1 & " " & Label3 & " " & Val(Text1) & "')"
        conexion_BD.Execute CHEp
        
        PROVp = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & List6.List(i) & "'," & Val(List5.List(i)) & ",'" & DTPicker4 & "','" & Label10 & "','CHEQUE CH')"
        conexion_BD.Execute PROVp
        Next
        
    Case Else 'CHEQUES PROPIOS Y TERCEROS
        For i = 0 To List10.ListCount - 1 ' BUCLE PARA LOS TERCEROS
        
            CHE3 = "insert into mov_cheques values (" & Val(List10.List(i)) & "," & Val(List9.List(i)) & ",'" & DTPicker4 & "','" & List1.List(i) & "','" & 0 & "','" & Combo1 & "','" & List3.List(i) & "','" & Label6 & " " & Combo1 & " " & Label3 & " " & Val(Text1) & "')"
            conexion_BD.Execute CHE3
            PROV3 = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & List3.List(i) & "'," & Val(List9.List(i)) & ",'" & DTPicker4 & "','" & Label10 & "','CHEQUE VT')"
            conexion_BD.Execute PROV3
            Label26 = Val(List3.List(i)) * -1
            L = "update entracheque set importe='" & Label26 & "' where n_interno =" & Val(List10.List(i)) & ""
            conexion_BD.Execute L
            
        Next
        
        For i = 0 To List4.ListCount - 1 'BUCLE PARA LOS PROPIOS
            prove = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & List6.List(i) & "'," & Val(List5.List(i)) & ",'" & DTPicker4 & "','" & Label10 & "','CHEQUE CH')"
            conexion_BD.Execute prove
            CHEp = "insert into mov_cheques values (" & Val(List4.List(i)) & "," & Val(List5.List(i)) & ",'" & DTPicker4 & "','" & List8.List(i) & "','" & 0 & "','" & Combo1 & "','" & List6.List(i) & "','" & Label6 & " " & Combo1 & " " & Label3 & " " & Val(Text1) & "')"
            conexion_BD.Execute CHEp
            SALECH = "insert into salecheque values (" & Val(List5.List(i)) & ",'" & List8.List(i) & "','" & Text8 & "','" & List6.List(i) & "','" & List7.List(i) & "','" & Combo1 & "','" & DTPicker4 & "'," & Val(List4.List(i)) & "," & Val(0) & "," & Val(Label10) & ")"
            conexion_BD.Execute SALECH
        Next
         
    
    End Select

    Call IMPRIMIR

Else ' COMBO3 = "CHEQUE"
       
    txtmon = CDbl(Label20)
    Call CONVERTIR
    Label37 = txtmonl
    
    
    rint = "insert into remito_interno values (" & Val(Label10) & ",'" & DTPicker4 & "','" & Label22 & "','" & Combo1 & "','" & Text8 & "','" & Label20 & "','" & usua & "')"
    conexion_BD.Execute rint
    
    ret = "insert into retencion values ('" & DTPicker4 & "','" & Combo1 & "','" & Text10 & "'," & Val(Label10) & ")"
    conexion_BD.Execute ret
    
    'prove = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & text11 & "'," & Val(0) & ",'" & DTPicker4 & "','" & Label10 & "')"
    'conexion_BD.Execute prove


    Select Case Combo4.Text
        Case "TERCEROS"
        For i = 0 To List10.ListCount - 1
        CHE3 = "insert into mov_cheques values (" & Val(List10.List(i)) & "," & Val(List9.List(i)) & ",'" & DTPicker4 & "','" & List1.List(i) & "','" & 0 & "','" & Combo1 & "','" & List3.List(i) & "','" & Label6 & " " & Combo1 & " " & Label3 & " " & Val(Text1) & "')"
        conexion_BD.Execute CHE3
        
        PROV3 = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & List3.List(i) & "'," & Val(List9.List(i)) & ",'" & DTPicker4 & "','" & Label10 & "','CHEQUE VT')"
        conexion_BD.Execute PROV3
        Label26 = Val(List3.List(i)) * -1
        L = "update entracheque set importe='" & Label26 & "' where n_interno =" & List10.List(i) & ""
        conexion_BD.Execute L
        
        Next
    
        
    Case "PROPIOS"
        For i = 0 To List4.ListCount - 1
        SALECH = "insert into salecheque values (" & Val(List5.List(i)) & ",'" & List8.List(i) & "','" & Text8 & "','" & List6.List(i) & "','" & List7.List(i) & "','" & Combo1 & "','" & DTPicker4 & "'," & Val(List4.List(i)) & "," & Val(0) & "," & Val(Label10) & ")"
        conexion_BD.Execute SALECH
        
        CHEp = "insert into mov_cheques values (" & Val(List4.List(i)) & "," & Val(List5.List(i)) & ",'" & DTPicker4 & "','" & List8.List(i) & "','" & 0 & "','" & Combo1 & "','" & List6.List(i) & "','" & Label6 & " " & Combo1 & " " & Label3 & " " & Val(Text1) & "')"
        conexion_BD.Execute CHEp
        
        PROVp = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & List6.List(i) & "'," & Val(List5.List(i)) & ",'" & DTPicker4 & "','" & Label10 & "','CHEQUE CH')"
        conexion_BD.Execute PROVp
        Next
        
    Case Else 'CHEQUES PROPIOS Y TERCEROS
        For i = 0 To List10.ListCount - 1 ' BUCLE PARA LOS TERCEROS
        
            CHE3 = "insert into mov_cheques values (" & Val(List10.List(i)) & "," & Val(List9.List(i)) & ",'" & DTPicker4 & "','" & List1.List(i) & "','" & 0 & "','" & Combo1 & "','" & List3.List(i) & "','" & Label6 & " " & Combo1 & " " & Label3 & " " & Val(Text1) & "')"
            conexion_BD.Execute CHE3
            PROV3 = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & List3.List(i) & "'," & Val(List9.List(i)) & ",'" & DTPicker4 & "','" & Label10 & "','CHEQUE VT')"
            conexion_BD.Execute PROV3
            Label26 = Val(List3.List(i)) * -1
            L = "update entracheque set importe='" & Label26 & "' where n_interno =" & Val(List10.List(i)) & ""
            conexion_BD.Execute L
            
        Next
        
        For i = 0 To List4.ListCount - 1 'BUCLE PARA LOS PROPIOS
            PROVp = " insert into mov_proveedor values ('" & Combo1 & "','" & Text8 & "','" & List6.List(i) & "'," & Val(List5.List(i)) & ",'" & DTPicker4 & "','" & Label10 & "','CHEQUE CH')"
            conexion_BD.Execute PROVp
            CHEp = "insert into mov_cheques values (" & Val(List4.List(i)) & "," & Val(List5.List(i)) & ",'" & DTPicker4 & "','" & List8.List(i) & "','" & 0 & "','" & Combo1 & "','" & List6.List(i) & "','" & Label6 & " " & Combo1 & " " & Label3 & " " & Val(Text1) & "')"
            conexion_BD.Execute CHEp
            SALECH = "insert into salecheque values (" & Val(List5.List(i)) & ",'" & List8.List(i) & "','" & Text8 & "','" & List6.List(i) & "','" & List7.List(i) & "','" & Combo1 & "','" & DTPicker4 & "'," & Val(List4.List(i)) & "," & Val(0) & "," & Val(Label10) & ")"
            conexion_BD.Execute SALECH
        Next
         
    
    End Select


End If
Command2.Visible = True

Call IMPRIMIR

Call NO
Mov_proveedores.Refresh

Frame1.Visible = False
Frame2.Visible = False
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
List9.Clear
List10.Clear
Text1 = ""
Text2 = ""
Text11 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
'Text8 = ""
Text9 = ""
Text10 = ""
Combo5.Clear
Else
MsgBox " Los montos no coiciden"
End If

End Sub

Private Sub Command6_Click()

If Text5 = Val(List5.List(i)) Then
    MsgBox " El Cheque ya se encuentra en la lista"
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text5.SetFocus
Else
    If DTPicker1 > minimo And DTPicker1 < maximo Then
        List4.AddItem (Label19)
        List5.AddItem (Text5)
        List6.AddItem (Text6)
        List7.AddItem (Text7)
        List8.AddItem (DTPicker1)

        m = MsgBox("Desea cargar otro cheque?", vbYesNo, "VIVERO SAN NICOLAS")

        If m = vbYes Then
            Label19 = Label19 + 1
            Text5 = ""
            Text6 = ""
            Text7 = ""
            Text5.SetFocus
            Command5.Visible = True
        Else
            Text5 = ""
            Text6 = ""
            Text7 = ""
            Command5.Visible = True
        End If
    Else
        MsgBox ("La fecha de vencimiento se encuentra fuera del rango")
    End If
End If
End Sub



Private Sub Command7_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
List9.Clear
List10.Clear
Text1 = ""
Text2 = ""
Text9 = ""
Text4 = ""
Text5 = ""
Label20 = ""


End Sub

Private Sub Command8_Click()
    'DOBLE = CDbl(Text9) '+ Val(Text4) + Val(Text6)
    'Label20 = DOBLE

validar = " select * from prove_a_pagar where factura= '" & Text1 & "'"
TABLA.Open validar, conexion_BD
Do While Not TABLA.EOF = True
Label37 = TABLA!factura
Label26 = TABLA!proveedor
TABLA.MoveNext
Loop
TABLA.Close
If Text1.Text = Label37 And Combo1.Text = Label26 Then
    MsgBox " La factura ya ha sido ingresada. Por favor verificar el número. " & vbNewLine & " IMPORTANTE: si es parte de la misma factura," & vbNewLine & " se puede colocar de esta manera:" & vbNewLine & " EJ: 1234(1)"
    Text1.SetFocus
Else
    PROVEaPAGAR = "insert into prove_a_pagar values ('" & Combo1 & "','" & Text1 & "','" & Text2 & "','" & Text9 & "','" & DTPicker4 & "','" & DTPicker2 & "','" & DTPicker3 & "'," & 0 & ")"
    conexion_BD.Execute PROVEaPAGAR
    
    fact = "insert into facturas values ('" & Text1 & "','" & Combo1 & "','" & Text2 & "','" & Text9 & "','" & DTPicker2 & "','" & DTPicker3 & "','" & Text9 & "','IMPAGO','0','0')"
    conexion_BD.Execute fact
    
    Text1 = ""
    Text2 = ""
    Text9 = ""
    Label20 = ""
    Frame3.Visible = False
    
End If
End Sub


Private Sub Command9_Click()
datos = Combo1.Text
sValor = "PROVEEDORES"
Detalles_ctas.Show
End Sub

Private Sub Form_Load()
DTPicker1 = Date
minimo = CDate(Me.DTPicker1.value) - 25
maximo = CDate(Me.DTPicker1.value) + 360

DTPicker2 = Date
DTPicker3 = Date
DTPicker4 = Date
Label22 = SISTEMA

Label9 = Date
d = "select * from proveedores order by nombre_prove"
TABLA.Open d, conexion_BD
Do While Not TABLA.EOF
    Combo1.AddItem TABLA!nombre_prove
    TABLA.MoveNext
Loop
TABLA.Close
End Sub

Private Sub Combo3_Click()
Select Case Combo3.Text
    Case "CHEQUE"
        
        Label4.Visible = True
        Combo4.Visible = True
        Label5.Visible = False
        Text4.Visible = False
        Command1.Visible = False
        
        
    Case "EFECTIVO"
       
        Label5.Visible = False
        Text4.Visible = False
        Label4.Visible = False
        Combo4.Visible = False
        Command1.Visible = True
        Command6.Visible = False
        Frame1.Visible = False
        Frame2.Visible = False
        
        
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
    
    Case "TARJETA"
       
        Label5.Visible = False
        Text4.Visible = False
        Label4.Visible = False
        Combo4.Visible = False
        Command1.Visible = True
        Command6.Visible = False
        Frame1.Visible = False
        Frame2.Visible = False
        
        
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
        
    Case "INSUMOS"
        ''' SE UTILIZA PARA LOS CDO SE PAGA POR EJEMPLO CON GASOIL
        Label5.Visible = False
        Text4.Visible = False
        Label4.Visible = False
        Combo4.Visible = False
        Command1.Visible = True
        Command6.Visible = False
        Frame1.Visible = False
        Frame2.Visible = False
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
        
        Label4.Visible = True
        Label5.Visible = True
        Label7.Visible = True
        Text4.Visible = True
        Combo4.Visible = True
        Command1.Visible = False
        
End Select
End Sub

Private Sub Combo4_Click()
Select Case Combo4.Text
    Case "PROPIOS"
        Frame1.Visible = True
        Frame2.Visible = False
        Label14.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label17.Visible = True
        Label18.Visible = True
        Label19.Visible = True
        'Label22.Visible = False
        'Combo5.Visible = False
        Text5.Visible = True
        Text6.Visible = True
        Text7.Visible = True
        DTPicker1.Visible = True
        Command6.Visible = True
    
        a = "select max(n_interno) from salecheque"
        TABLA.Open a, conexion_BD
        Label19 = TABLA.Fields(0) + 1
        TABLA.Close
    Case "TERCEROS"
        Frame2.Visible = True
        Frame1.Visible = False
        'Label22.Visible = True
        'Combo5.Visible = True
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
        Command6.Visible = False
        
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
    Case Else
        Frame1.Visible = True
        Frame2.Visible = True
        Frame4.Visible = True
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
        Command6.Visible = True
        
        'Label22.Visible = True
        'Combo5.Visible = True
        
        a = "select max(n_interno) from salecheque"
        TABLA.Open a, conexion_BD
        Label19 = TABLA.Fields(0) + 1
        TABLA.Close
        
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
Private Sub VALIDAR_SUMA()
For i = 0 To List4.ListCount - 1
    If CDbl(List6.List(i)) = "0" Then
        propios = 0
    Else
    propios = CDbl(propios) + CDbl(List6.List(i))
    End If
Next

For i = 0 To List10.ListCount - 1
    If CDbl(List10.List(i)) = "0" Then
        terceros = 0
    Else
        terceros = CDbl(terceros) + CDbl(List3.List(i))
    End If
Next
suma = CDbl(propios) + CDbl(terceros)

If Text10 = "" Then
    Text10 = 0
End If
If Text4 = "" Then
    Text4 = 0
End If
suma1 = CDbl(Text10) + CDbl(Text4)
valido = CDbl(suma) + CDbl(suma1)

End Sub
Private Sub IMPRIMIR()

Printer.CurrentX = 30
Printer.CurrentY = 80
Printer.FontSize = 10
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); "Nº de recibo: "; Label10; 'label 21= n_interno
Printer.Print Tab(110); DTPicker4
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí/mos de "; SISTEMA
'Printer.Print Tab(15); " Recibí/mos de STELLA DAVIRE"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos $ "; Text11
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " (en letras:  "; Label37; " )"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " En concepto de: "; Text2; Text8; '" por la factura Nº: "; Text1
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Según el siguiente detalle:"
Printer.Print Tab(15); ""
Printer.Print Tab(15); "En efectivo: $ "; Text4; " En retenciones: $ "; Text10;
Printer.Print Tab(15); ""
Printer.Print Tab(15); "INTERNO Nº "; Tab(35); "CHEQUE Nº"; Tab(55); "IMPORTE"; Tab(75); "BANCO"; Tab(98); "VENCIMIENTO"; ' Tab(95); '"CUIT:"; Tab(100)
Printer.Print Tab(15); "----------------------------------------------------------------------------------------------------------------------------------"
tot = 0
tot2 = 0

For a = 0 To List4.ListCount - 1
    ' LISTADO DE CHEQUES PROPIOS
    List4.ListIndex = a: n_interno = List4.List(a)
    List5.ListIndex = a: n_cheque = List5.List(a)
    List6.ListIndex = a: importe = List6.List(a)
    List7.ListIndex = a: banco = List7.List(a)
    List8.ListIndex = a: fecha = List8.List(a)

Printer.Print Tab(15); " CH "; n_interno; Tab(35); n_cheque; Tab(55); importe; Tab(75); banco; Tab(98); fecha; 'Tab(95); 'CUIT; Tab(100)
Next

For b = 0 To List10.ListCount - 1
    'LISTADO DE CHEQUES DE TERCEROS
    
    List10.ListIndex = b: n_internoB = List10.List(b)
    List9.ListIndex = b: n_chequeB = List9.List(b)
    List3.ListIndex = b: importeB = List3.List(b)
    List2.ListIndex = b: bancoB = List2.List(b)
    List1.ListIndex = b: fechaB = List1.List(b)
Printer.Print Tab(15); " VT "; n_internoB; " (T)"; Tab(35); n_chequeB; Tab(55); importeB; Tab(75); bancoB; Tab(98); fechaB; ' Tab(85); 'CUIT; Tab(100)
Next

Printer.Print Tab(15); "----------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Monto total en $ "; Text11
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(70); Combo1.Text
Printer.Print Tab(15); ""
Printer.Print Tab(15); "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ";

'''''''''''''''''''
'' IMPRIME COPIA ''
'''''''''''''''''''
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); "Nº de recibo: "; Label10; 'label 21= n_interno
Printer.Print Tab(110); DTPicker4
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí/mos de "; SISTEMA
'Printer.Print Tab(15); " Recibí/mos de STELLA DAVIRE"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos $ "; Text11
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " ( "; Label37; " )"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " En concepto de: "; Text2; Text8; '" por la factura Nº: "; Text1
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Según el siguiente detalle:"
Printer.Print Tab(15); ""
Printer.Print Tab(15); "En efectivo: $ "; Text4; " En retenciones: $ "; Text10;
Printer.Print Tab(15); ""
Printer.Print Tab(15); "INTERNO Nº "; Tab(35); "CHEQUE Nº"; Tab(55); "IMPORTE"; Tab(75); "BANCO"; Tab(98); "VENCIMIENTO"; ' Tab(95); '"CUIT:"; Tab(100)
Printer.Print Tab(15); "----------------------------------------------------------------------------------------------------------------------------------"
tot = 0
tot2 = 0

For a = 0 To List4.ListCount - 1
    ' LISTADO DE CHEQUES PROPIOS
    List4.ListIndex = a: n_interno = List4.List(a)
    List5.ListIndex = a: n_cheque = List5.List(a)
    List6.ListIndex = a: importe = List6.List(a)
    List7.ListIndex = a: banco = List7.List(a)
    List8.ListIndex = a: fecha = List8.List(a)

Printer.Print Tab(15); " CH "; n_interno; Tab(35); n_cheque; Tab(55); importe; Tab(75); banco; Tab(98); fecha; 'Tab(95); 'CUIT; Tab(100)
Next

For b = 0 To List10.ListCount - 1
    'LISTADO DE CHEQUES DE TERCEROS
    
    List10.ListIndex = b: n_internoB = List10.List(b)
    List9.ListIndex = b: n_chequeB = List9.List(b)
    List3.ListIndex = b: importeB = List3.List(b)
    List2.ListIndex = b: bancoB = List2.List(b)
    List1.ListIndex = b: fechaB = List1.List(b)
Printer.Print Tab(15); " VT "; n_internoB; Tab(35); n_chequeB; Tab(55); importeB; Tab(75); bancoB; Tab(98); fechaB; ' Tab(85); 'CUIT; Tab(100)
Next

Printer.Print Tab(15); "----------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Monto total en $ "; Text11
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(70); Combo1.Text
Printer.Print Tab(15); ""
Printer.Print Tab(15); "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ";


Printer.EndDoc
End Sub

Private Sub NO()
       
    Text9 = "0"
    Text11 = "0"
    Text4 = "0"
    Text5 = "0"
    Text6 = "0"
    Text7 = "0"
    Label19 = "0"
End Sub



Private Sub MSFlexGrid1_Click()
interno = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
cheque = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
importe = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
banco = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
vto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4)

List10.AddItem interno
List9.AddItem cheque
List3.AddItem importe
List2.AddItem banco
List1.AddItem vto

r = MsgBox("Desea cargar otro cheque?", vbYesNo)
If r = vbYes Then
    MSFlexGrid1.RemoveItem (MSFlexGrid1.RowSel)
    'If List10.List(i) = interno Then
    'MsgBox "El cheque ya ha sido ingresado"
    'End If
    Frame4.Visible = True
Else
    Command5.Visible = True
    Frame4.Visible = False
End If
End Sub

Private Sub Text1_GotFocus()
If Text1 = "0" Then
    Text1 = ""
End If
End Sub



Private Sub Text10_GotFocus()
If Text10 = "0" Then
    Text10 = ""
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text11_GotFocus()
If Text11 = "0" Then
    Text11 = ""
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text2_GotFocus()
If Text2 = "0" Then
    Text2 = ""
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


Private Sub Text7_GotFocus()
If Text7 = "0" Then
    Text7 = ""
End If
End Sub

Private Sub Text9_GotFocus()
If Text9 = "0" Then
    Text9 = ""
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub



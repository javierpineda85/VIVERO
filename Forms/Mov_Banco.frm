VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mov_Banco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos del Banco"
   ClientHeight    =   9360
   ClientLeft      =   2925
   ClientTop       =   960
   ClientWidth     =   8115
   Icon            =   "Mov_Banco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   8115
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Cheques ingresados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   28
      Top             =   5880
      Visible         =   0   'False
      Width           =   5295
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
         Height          =   1500
         Left            =   3960
         TabIndex        =   33
         Top             =   840
         Width           =   1095
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
         Height          =   1500
         Left            =   3000
         TabIndex        =   32
         Top             =   840
         Width           =   975
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
         Height          =   1500
         Left            =   2040
         TabIndex        =   31
         Top             =   840
         Width           =   975
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
         Height          =   1500
         Left            =   1080
         TabIndex        =   30
         Top             =   840
         Width           =   975
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
         Height          =   1500
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label19 
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
         Left            =   360
         TabIndex        =   38
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label18 
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
         TabIndex        =   37
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label17 
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
         Left            =   2160
         TabIndex        =   36
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label16 
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
         Left            =   3120
         TabIndex        =   35
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Venc."
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
         Left            =   4080
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4215
      Left            =   3960
      TabIndex        =   43
      Top             =   2880
      Visible         =   0   'False
      Width           =   3855
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   1680
         TabIndex        =   55
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
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
         CurrentDate     =   41093
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
         Left            =   1680
         TabIndex        =   54
         Top             =   2760
         Width           =   1815
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
         Left            =   1680
         TabIndex        =   53
         Top             =   2160
         Width           =   1815
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
         Left            =   1680
         TabIndex        =   52
         Top             =   960
         Width           =   1815
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
         TabIndex        =   51
         Top             =   360
         Width           =   1815
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
         Height          =   735
         Left            =   2400
         Picture         =   "Mov_Banco.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ingresar"
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
         Left            =   1320
         Picture         =   "Mov_Banco.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label21 
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
         Height          =   375
         Left            =   360
         TabIndex        =   49
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Importe: $"
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
         TabIndex        =   48
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Fecha vto:"
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
         TabIndex        =   47
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label28 
         Caption         =   "Nº cheque:"
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
         TabIndex        =   46
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "Nº  interno:"
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
         TabIndex        =   45
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   240
      TabIndex        =   39
      Top             =   2880
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton Command8 
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
         Left            =   5160
         Picture         =   "Mov_Banco.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4920
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4695
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   8281
         _Version        =   393216
         BackColor       =   16777152
         SelectionMode   =   1
         AllowUserResizing=   2
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
      Begin VB.CommandButton Command5 
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
         Left            =   6240
         Picture         =   "Mov_Banco.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4920
         Width           =   975
      End
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
      ItemData        =   "Mov_Banco.frx":1BB2
      Left            =   1800
      List            =   "Mov_Banco.frx":1BBC
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   61472769
      CurrentDate     =   41085
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
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
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
      Left            =   1560
      TabIndex        =   5
      Text            =   "0"
      Top             =   3840
      Width           =   1215
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
      TabIndex        =   6
      Text            =   "0"
      Top             =   3840
      Width           =   1215
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
      Left            =   360
      Picture         =   "Mov_Banco.frx":1BD3
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
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
      Left            =   1560
      Picture         =   "Mov_Banco.frx":215D
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5160
      TabIndex        =   13
      Top             =   8760
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
      ItemData        =   "Mov_Banco.frx":26E7
      Left            =   1440
      List            =   "Mov_Banco.frx":26F1
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "x fecha"
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
      Picture         =   "Mov_Banco.frx":270B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Devolucion"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2655
      Left            =   240
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4683
      _Version        =   393216
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   61472769
      CurrentDate     =   41023
      MinDate         =   2
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   4560
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   61472769
      CurrentDate     =   41023
      MinDate         =   2
   End
   Begin VB.Label Label20 
      Caption         =   "VIVERO"
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "L11= importe*-1"
      Height          =   375
      Left            =   4680
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheques:"
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
      Left            =   360
      TabIndex        =   27
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha de Mov.:"
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
      Left            =   3720
      TabIndex        =   26
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos del Banco"
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
      Left            =   1680
      TabIndex        =   25
      Top             =   480
      Width           =   5055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha L2"
      Height          =   255
      Left            =   6600
      TabIndex        =   24
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle:"
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
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto: $"
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
      TabIndex        =   22
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo: $"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Listado desde:"
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
      TabIndex        =   20
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
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
      Left            =   4080
      TabIndex        =   19
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
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
      Left            =   4320
      TabIndex        =   18
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
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
      TabIndex        =   17
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "L13= nremito"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Mov_Banco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo2_Click()
If Combo2 = "PROPIOS" Then
    Frame3.Visible = True
    Frame2.Visible = False
    a = "select max(n_interno) from salecheque"
    TABLA.Open a, conexion_BD
    Text5 = TABLA.Fields(0) + 1
    TABLA.Close
Else
    Frame2.Visible = True
    Frame3.Visible = False
    MSFlexGrid2.Clear
    Call AltaCheque
    
    P = "select * from entracheque order by fecha_vto"
    TABLA.Open P, conexion_BD
    
    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!n_interno
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!n_cheque
        MSFlexGrid2.TextMatrix(lin, 2) = TABLA!importe
        MSFlexGrid2.TextMatrix(lin, 3) = TABLA!banco
        MSFlexGrid2.TextMatrix(lin, 4) = TABLA!fecha_vto

   
        If TABLA!importe <= "0" Or TABLA!rechazado = "-1" Then
            MSFlexGrid2.Row = lin
            MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
            lin = lin - 1
        Else
    
            importe = CDbl(importe) + TABLA!importe
        
        End If
    
        TABLA.MoveNext
    Loop
    'Label3 = Format(importe, "currency")
    TABLA.Close
End If
End Sub

Private Sub Command1_Click()
Dim valores As String
If Text1 = "" Or Combo1 = "" Then
    MsgBox "Debe cargar el detalle y el Destinatario", vbOKOnly, "VIVERO SAN NICOLAS"
    Text1.SetFocus
    If Text2 = "" Then
        Text2 = "0"
        txtmon = Text3
    Else
        Text3 = "0"
        txtmon = Text2
    End If
Else
    'txtmon = cdbl(Text3)
    'Call CONVERTIR
    'Label14 = txtmonl
    Label14 = CDbl(Text2) + CDbl(Text3)
    'Call IMPRIMIR

    remito = "select max(n_remito) from remito_interno"
    TABLA.Open remito, conexion_BD
    Label13 = TABLA.Fields(0) + 1
    TABLA.Close
    
    If Combo1 = "Deposito" Then
        
        rint = "insert into remito_interno values (" & Val(Label13) & ",'" & DTPicker3 & "','VIVERO SAN NICOLAS S.A.','" & Combo2 & "','" & Text1 & "'," & Val(Label14) & ",'" & usua & "')"
        conexion_BD.Execute rint
    
        'GUARDAMOS LOS DATOS EN EL BANCO
        bandep = " insert into banco values ('" & DTPicker3 & "','" & Text1 & "', '" & Label14 & "','" & 0 & "','" & Label13 & "','" & usua & "')"
        conexion_BD.Execute bandep
        
        'A LA CAJA VA SOLO LO QUE SE DEPOSITE EN EFECTIVO
        acaja = "insert into mov_caja values ('" & DTPicker3 & "','" & Text1 & "','" & 0 & "','" & Text3 & "','" & Label13 & "')"
        conexion_BD.Execute acaja
        
        If Combo2 = "PROPIOS" Then
            For i = 0 To List1.ListCount - 1
                chepro = "insert into salecheque values (" & Val(List2.List(i)) & ",'" & List5.List(i) & "','" & Text1 & "','" & _
                    List3.List(i) & "','" & List4.List(i) & "','" & Label20 & "','" & DTPicker3 & "'," & Val(List1.List(i)) & "," & 0 & "," & Val(Label13) & " )"
                    conexion_BD.Execute chepro
                
                enbanco = "insert into enbanco values (" & Val(List2.List(i)) & "," & Val(List1.List(i)) & "," & _
                    Val(List5.List(i)) & "," & Val(List3.List(i)) * -1 & ",'" & DTPicker3 & "')"
                    conexion_BD.Execute enbanco
                    
                movcheques = "insert into mov_cheques values (" & Val(List1.List(i)) & "," & Val(List2.List(i)) & ",'" & DTPicker3 & "','" & 0 & "','" & _
                Combo1 & "','DEPOSITO','" & List3.List(i) & "','" & Text1 & "')"
                    conexion_BD.Execute movcheques
                Next
        Else
            For i = 0 To List1.ListCount - 1

                Label11 = Val(List3.List(i)) * -1
                L = "update entracheque set importe='" & Label11 & "' where n_interno =" & List1.List(i) & ""
                conexion_BD.Execute L
                
                enbanco = "insert into enbanco values (" & Val(List2.List(i)) & "," & Val(List1.List(i)) & "," & _
                    Val(List5.List(i)) & "," & Val(List3.List(i)) * -1 & ",'" & DTPicker3 & "')"
                    conexion_BD.Execute enbanco
                    
                movcheques = "insert into mov_cheques values (" & Val(List1.List(i)) & "," & Val(List2.List(i)) & ",'" & DTPicker3 & "','" & 0 & "','" & _
                    Combo1 & "','DEPOSITO ','" & List3.List(i) & "','" & Text1 & "')"
                    conexion_BD.Execute movcheques
            Next
        End If

    Else
        ' combo1="extraccion"
        
        rint = "insert into remito_interno values (" & Val(Label13) & ",'" & DTPicker3 & "','VIVERO SAN NICOLAS S.A.','" & Combo2 & "','" & Text1 & "'," & Val(Label14) & ",'" & usua & "')"
        conexion_BD.Execute rint
        
        'Label14 = Text2 * -1
        bandep = " insert into banco values ('" & DTPicker3 & "','" & Text1 & "', '" & 0 & "','" & Label14 & "','" & Label13 & "','" & usua & "')"
        conexion_BD.Execute bandep
        
        acaja = "insert into mov_caja values ('" & DTPicker3 & "','" & Text1 & "','" & Label14 & "','" & 0 & "','" & Label13 & "')"
        conexion_BD.Execute acaja
        
        'MODIFICAMOS LOS DATOS DE LOS CHEQUES
        'aca multiplicamos el list3 *-1
        
        If Combo2 = "PROPIOS" Then
            For i = 0 To List1.ListCount - 1
                chepro = "insert into salecheque values (" & Val(List2.List(i)) & ",'" & List5.List(i) & "','" & Text1 & "','" & List3.List(i) * -1 & "','" & List4.List(i) & "','" & Label20 & "','" & DTPicker3 & "'," & Val(List1.List(i)) & "," & Val(0) & "," & Val(Label13) & ")"
                    conexion_BD.Execute chepro
                
                enbanco = "insert into enbanco values (" & Val(List2.List(i)) & "," & Val(List1.List(i)) & "," & _
                    Val(List5.List(i)) & "," & Val(List3.List(i)) * -1 & ",'" & DTPicker3 & "')"
                    conexion_BD.Execute enbanco
                    
                movcheques = "insert into mov_cheques values (" & Val(List1.List(i)) & "," & Val(List2.List(i)) & ",'" & DTPicker3 & "','" & 0 & "','" & _
                    Combo1 & "','EXTRACCION','" & List3.List(i) & "','" & Text1 & "')"
                    conexion_BD.Execute movcheques
                Next
        Else
            For i = 0 To List1.ListCount - 1
                Label11 = Val(List3.List(i)) * -1
                L = "update entracheque set importe='" & Label11 & "' where n_interno =" & List1.List(i) & ""
                conexion_BD.Execute L
                
                enbanco = "insert into enbanco values (" & Val(List2.List(i)) & "," & Val(List1.List(i)) & "," & _
                    Val(List5.List(i)) & "," & Val(List3.List(i)) * -1 & ",'" & DTPicker3 & "')"
                    conexion_BD.Execute enbanco
                    
                movcheques = "insert into mov_cheques values (" & Val(List1.List(i)) & "," & Val(List2.List(i)) & ",'" & DTPicker3 & "','" & 0 & "','" & _
                    Combo1 & "','EXTRACCION','" & List3.List(i) & "','" & Text1 & "')"
                    conexion_BD.Execute movcheques
            Next
        End If
            
    End If
    Call IMPRIMIR
    
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear

End If



Label11 = ""
Label12 = ""
Label13 = ""
Text1 = ""
Text2 = "0"
Text3 = "0"
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""


End Sub

Private Sub Command2_Click()
LISTADO = "select * from banco order by fecha DESC" ' where fecha between # " & DTPicker1 & "# and # " & DTPicker2 & "# order by fecha"
TABLA.Open LISTADO, conexion_BD
Call ALTAGRID
MSFlexGrid1.Visible = True
Label7.Visible = True
saldo = 0
entra = 0
sale = 0

    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!detalle
        MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!ingreso, "currency")
        MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!egreso, "currency")

        entra = CDbl(entra) + TABLA!ingreso
        sale = CDbl(sale) + TABLA!egreso
        TABLA.MoveNext
    Loop
TABLA.Close
saldo = CDbl(entra) - CDbl(sale)
If saldo < 0 Then
    Text4.BackColor = &HFF&
    Text4.ForeColor = &HFFFFFF
    
Else
    Text4.BackColor = vbGreen
    Text4.ForeColor = vbBlack
    
End If
Text4 = Format(saldo, "currency")

End Sub

Private Sub Command3_Click()
Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)

LISTADO = "select * from banco where fecha >= " + CStr(inicio) + " and fecha <= " + CStr(final) + " order by fecha"
TABLA.Open LISTADO, conexion_BD
Call ALTAGRID
MSFlexGrid1.Visible = True
Label7.Visible = True

saldo = 0
entra = 0
sale = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!detalle
        MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!ingreso, "currency")
        MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!egreso, "currency")
        entra = CDbl(entra) + TABLA!ingreso
        sale = CDbl(sale) + TABLA!egreso
        TABLA.MoveNext
    Loop
TABLA.Close
'saldo = CDbl(entra) - CDbl(sale)
'If saldo < 0 Then
'    Text4.BackColor = &HFF&
'    Text4.ForeColor = &HFFFFFF
'
'Else
'    Text4.BackColor = vbGreen
'    Text4.ForeColor = vbBlack
'
'End If
'Text4 = Format(saldo, "currency")

End Sub

Private Sub Command5_Click()
interno = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 0)
ncheque = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 1)
importe = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 2)
banco = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 3)
venc = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 4)

If List1.List(i) = interno Then
    MsgBox "El Cheque ya ha sido ingresado"
Else
    List1.AddItem interno
    List2.AddItem ncheque
    List3.AddItem importe
    List4.AddItem banco
    List5.AddItem venc
    Text3 = 0
    Text2 = 0
    MSFlexGrid2.RemoveItem (MSFlexGrid2.RowSel)
    For i = 0 To List3.ListCount - 1
        Text2 = CDbl(Text2) + CDbl(List3.List(i))
    Next
    
    res = MsgBox("desea cargar otro cheque?", vbYesNo)
    If res = vbYes Then
        Frame2.Visible = True
        Frame1.Visible = False
    Else
        Frame2.Visible = False
        Frame1.Visible = True
    End If
End If
End Sub

Private Sub Command6_Click()
Frame1.Visible = True
interno = Text5
cheque = Text6
vto = DTPicker4
importe = Text7
banco = Text8
If List1.List(i) = interno Then
    MsgBox "El Cheque ya ha sido ingresado"
Else
    List1.AddItem interno
    List2.AddItem cheque
    List3.AddItem importe
    List4.AddItem banco
    List5.AddItem vto
    
    For i = 0 To List3.ListCount - 1
        Text2 = CDbl(Text2) + CDbl(List3.List(i))
    Next
    
    res = MsgBox("desea cargar otro cheque?", vbYesNo)
    
    If res = vbYes Then
        Frame3.Visible = True
        Text5 = Text5 + 1
        Text6 = ""
        Text7 = ""
        Text8 = ""
               
    Else
        Frame3.Visible = False
        Text5 = ""
        Text6 = ""
        Text7 = ""
        Text8 = ""
    End If
End If

End Sub

Private Sub Command7_Click()
Frame3.Visible = False
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
End Sub

Private Sub Command8_Click()
Frame2.Visible = False
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
DTPicker4 = Date
Label2 = Date


End Sub
Private Sub ALTAGRID()
MSFlexGrid1.FixedCols = 0
MSFlexGrid1.Cols = 4
MSFlexGrid1.FixedRows = 1
MSFlexGrid1.Rows = 2
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "FECHA"
MSFlexGrid1.TextMatrix(0, 1) = "DETALLE"
MSFlexGrid1.TextMatrix(0, 2) = "INGRESO"
MSFlexGrid1.TextMatrix(0, 3) = "EGRESO"

MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 1500


End Sub
Private Sub AltaCheque()
With MSFlexGrid2
.FixedCols = 0
.Cols = 5
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "Nº INTERNO"
.TextMatrix(0, 1) = "Nº DE CHEQUE"
.TextMatrix(0, 2) = "IMPORTE"
.TextMatrix(0, 3) = "BANCO"
.TextMatrix(0, 4) = "FECHA DE VTO"

.ColWidth(0) = 1000
.ColWidth(1) = 1700
.ColWidth(2) = 1500
.ColWidth(3) = 2000
.ColWidth(4) = 1500

End With
End Sub

Private Sub IMPRIMIR()


'remito = "select max(n_remito) from remito_interno"
'TABLA.Open remito, conexion_BD
'Label13 = TABLA.Fields(0) + 1
'TABLA.Close

Printer.CurrentX = 30
Printer.CurrentY = 80
Printer.FontSize = 10
Printer.PaperSize = 9 ' papel A4

Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); " Remito Nº: "; Label13
Printer.Print Tab(98); DTPicker3
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " GESTION BANCARIA ";
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Se genero un movimiento bancario por la suma de: $ "; Label14;
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " ( "; Label14; " ) "
Printer.Print Tab(15); ""
Printer.Print Tab(15); " en concepto de "; Combo1.Text; ". "; Text1
Printer.Print Tab(15); ""
Printer.Print Tab(15); "INTERNO Nº "; Tab(35); "CHEQUE Nº"; Tab(55); "IMPORTE"; Tab(75); "BANCO"; Tab(98); "VENCIMIENTO"; 'Tab(95); "CUIT:"; Tab(110)
Printer.Print Tab(15); "=============================================================================="

For a = 0 To List1.ListCount - 1
    List1.ListIndex = a: n_interno = List1.List(a)
    List2.ListIndex = a: n_cheque = List2.List(a)
    List3.ListIndex = a: importe = List3.List(a)
    List4.ListIndex = a: banco = List4.List(a)
    List5.ListIndex = a: fecha = List5.List(a)

Printer.Print Tab(15); " CH "; n_interno; Tab(35); n_cheque; Tab(55); importe; Tab(75); banco; Tab(98); fecha; 'Tab(95); CUIT; Tab(110)

Next

Printer.Print Tab(15); ""
Printer.Print Tab(15); " Son  $ "; Label14
'Printer.Print Tab(60); "_____________";
'Printer.Print Tab(63); Combo2.Text
Printer.Print Tab(15); ""
Printer.Print Tab(15); "------------------------------------------------------------------------------------------------------------------------------------------------------";

Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); " Remito Nº: "; Label13
Printer.Print Tab(98); DTPicker3
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " GESTION BANCARIA ";
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Se genero un movimiento bancario por la suma de: $ "; Label14;
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " ( "; Label14; " ) "
Printer.Print Tab(15); ""
Printer.Print Tab(15); " en concepto de "; Combo1.Text; ". "; Text1
Printer.Print Tab(15); ""
Printer.Print Tab(15); "INTERNO Nº "; Tab(35); "CHEQUE Nº"; Tab(55); "IMPORTE"; Tab(75); "BANCO"; Tab(98); "VENCIMIENTO"; 'Tab(95); "CUIT:"; Tab(110)
Printer.Print Tab(15); "=============================================================================="

For a = 0 To List1.ListCount - 1
    List1.ListIndex = a: n_interno = List1.List(a)
    List2.ListIndex = a: n_cheque = List2.List(a)
    List3.ListIndex = a: importe = List3.List(a)
    List4.ListIndex = a: banco = List4.List(a)
    List5.ListIndex = a: fecha = List5.List(a)

Printer.Print Tab(15); " CH "; n_interno; Tab(35); n_cheque; Tab(55); importe; Tab(75); banco; Tab(98); fecha; 'Tab(95); CUIT; Tab(110)
Next
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Son  $ "; Label14
'Printer.Print Tab(60); "_____________";
'Printer.Print Tab(63); Combo2.Text
Printer.Print Tab(15); ""
Printer.Print Tab(15); "------------------------------------------------------------------------------------------------------------------------------------------------------";

Printer.EndDoc

End Sub



Private Sub Text2_GotFocus()
If Text2 = "0" Then
    Text2 = ""
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
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

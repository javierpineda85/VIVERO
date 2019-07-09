VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Listado_ingreso_egreso 
   Caption         =   "Listado de Ingresos y Egresos"
   ClientHeight    =   9285
   ClientLeft      =   1500
   ClientTop       =   1110
   ClientWidth     =   12645
   Icon            =   "Listado_ingreso_egreso.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   12645
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   11033
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Totales Generales"
      TabPicture(0)   =   "Listado_ingreso_egreso.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label7"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label8"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label9"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label10"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label11"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label12"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TotalGral"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label14"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label15"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label16"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label17"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label18"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label19"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label20"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label21"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label22"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label23"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label24"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label25"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label13"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label26"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label27"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Line17"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label28"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Line14"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Line15"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Line16"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Line18"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Line19"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label29"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "Ingresos"
      TabPicture(1)   =   "Listado_ingreso_egreso.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Egresos"
      TabPicture(2)   =   "Listado_ingreso_egreso.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSFlexGrid2"
      Tab(2).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   5055
         Left            =   -74520
         TabIndex        =   20
         Top             =   720
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8916
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5055
         Left            =   -73320
         TabIndex        =   19
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8916
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
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "Label29"
         Height          =   255
         Left            =   7920
         TabIndex        =   37
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Line Line19 
         X1              =   11640
         X2              =   11640
         Y1              =   3960
         Y2              =   5640
      End
      Begin VB.Line Line18 
         X1              =   7800
         X2              =   7800
         Y1              =   3960
         Y2              =   5640
      End
      Begin VB.Line Line16 
         X1              =   5880
         X2              =   5880
         Y1              =   3960
         Y2              =   5640
      End
      Begin VB.Line Line15 
         X1              =   2040
         X2              =   2040
         Y1              =   3960
         Y2              =   4800
      End
      Begin VB.Line Line14 
         X1              =   240
         X2              =   240
         Y1              =   3960
         Y2              =   4800
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "Label28"
         Height          =   255
         Left            =   4080
         TabIndex        =   36
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Line Line17 
         X1              =   5880
         X2              =   11640
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Label27"
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
         Left            =   7920
         TabIndex        =   35
         Top             =   4200
         Width           =   3615
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "Label26"
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
         Left            =   2160
         TabIndex        =   34
         Top             =   4200
         Width           =   3615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "SALDO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   33
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "Label25"
         Height          =   255
         Left            =   9840
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "Label24"
         Height          =   255
         Left            =   7920
         TabIndex        =   31
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Label23"
         Height          =   255
         Left            =   6000
         TabIndex        =   30
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Label22"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Label21"
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Label20"
         Height          =   255
         Left            =   9840
         TabIndex        =   27
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Label19"
         Height          =   255
         Left            =   7920
         TabIndex        =   26
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Label18"
         Height          =   255
         Left            =   6000
         TabIndex        =   25
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "TOTAL EGRESO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6000
         TabIndex        =   24
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Label16"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Label15"
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Label14"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label TotalGral 
         Alignment       =   2  'Center
         Caption         =   "TotalGral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   18
         Top             =   5040
         Width           =   3495
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "TOTAL INGRESO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "SUBTOTALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "EGRESO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "INGRESO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "RETENCION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9960
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "INSUMOS / TARJETA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   " CHEQUES PROPIOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "CHEQUES TERCEROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "EFECTIVO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Line Line13 
         X1              =   9720
         X2              =   9720
         Y1              =   720
         Y2              =   3120
      End
      Begin VB.Line Line12 
         X1              =   7800
         X2              =   7800
         Y1              =   720
         Y2              =   3120
      End
      Begin VB.Line Line11 
         X1              =   3960
         X2              =   3960
         Y1              =   720
         Y2              =   3120
      End
      Begin VB.Line Line10 
         X1              =   2040
         X2              =   2040
         Y1              =   720
         Y2              =   3120
      End
      Begin VB.Line Line9 
         X1              =   5880
         X2              =   5880
         Y1              =   720
         Y2              =   3120
      End
      Begin VB.Line Line8 
         X1              =   240
         X2              =   11640
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line7 
         X1              =   240
         X2              =   11640
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line6 
         X1              =   240
         X2              =   11640
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   11640
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line4 
         X1              =   11640
         X2              =   11640
         Y1              =   720
         Y2              =   3120
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   11640
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   240
         Y1              =   1440
         Y2              =   3120
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   11640
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   9375
      Begin VB.CommandButton Command3 
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
         Left            =   8040
         Picture         =   "Listado_ingreso_egreso.frx":05DE
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Listado_ingreso_egreso.frx":0B68
         DragIcon        =   "Listado_ingreso_egreso.frx":10F2
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
         Left            =   5640
         Picture         =   "Listado_ingreso_egreso.frx":167C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
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
         Left            =   6840
         Picture         =   "Listado_ingreso_egreso.frx":1C06
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   480
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
         CurrentDate     =   41185
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   480
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
         CurrentDate     =   40909
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Ingresos y Egresos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "Listado_ingreso_egreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call ALTA_INGRESO_FECHA
Call ALTA_EGRESO_FECHA

''' INGRESOS'''

With MSFlexGrid1

For i = 1 To .Rows - 1
    If .TextMatrix(i, 2) <> "EFECTIVO" Then
    If .TextMatrix(i, 2) = "" Or .TextMatrix(i, 3) = "" Then
        .TextMatrix(i, 2) = Format(0, "currency")
        .TextMatrix(i, 3) = Format(0, "currency")
        .TextMatrix(i, 4) = Format(0, "currency")
    Else
        efectivo_ing = CDbl(efectivo_ing) + .TextMatrix(i, 2)
        cheque_ing = CDbl(cheque_ing) + .TextMatrix(i, 3)
        insu_ing = CDbl(insu_ing) + .TextMatrix(i, 4)
        total = CDbl(.TextMatrix(i, 2)) + CDbl(.TextMatrix(i, 3)) + CDbl(.TextMatrix(i, 4))
        .TextMatrix(i, 1) = Format(total, "currency")
    End If
    End If
Next
Label14 = Format(efectivo_ing, "currency")
Label15 = Format(cheque_ing, "currency")
Label29 = Format(insu_ing, "currency")
'Label22 = Label15
End With

'''EGRESOS'''

With MSFlexGrid2

For i = 1 To .Rows - 1
    If .TextMatrix(i, 2) <> "EFECTIVO" Then
    If .TextMatrix(i, 2) = "" Or .TextMatrix(i, 3) = "" Then
        .TextMatrix(i, 2) = Format(0, "currency")
        .TextMatrix(i, 3) = Format(0, "currency")
        .TextMatrix(i, 4) = Format(0, "currency")
        .TextMatrix(i, 5) = Format(0, "currency")
    Else
        efectivo_sale = CDbl(efectivo_sale) + .TextMatrix(i, 2)
        cheque_pro = CDbl(cheque_pro) + .TextMatrix(i, 3)
        cheque_ter = CDbl(cheque_ter) + .TextMatrix(i, 4)
        reten = CDbl(reten) + .TextMatrix(i, 5)
        insumos = CDbl(insumos) + .TextMatrix(i, 6)
        total = CDbl(.TextMatrix(i, 2)) + CDbl(.TextMatrix(i, 3)) + CDbl(.TextMatrix(i, 4)) + CDbl(.TextMatrix(i, 5)) + CDbl(.TextMatrix(i, 6))
        .TextMatrix(i, 1) = Format(total, "currency")
        
    End If
    End If
Next

End With

Label16 = Format(efectivo_sale, "currency")
Label18 = Format(cheque_pro, "currency")
Label28 = Format(cheque_ter, "currency")
Label19 = Format(insumos, "currency")
Label20 = Format(reten, "currency")
'Label23 = Label18
'Label24 = Label19
'Label25 = Label20
'subefect = CDbl(Label14) - CDbl(Label16)
'Label21 = Format(subefect, "currency")

'TOTAL INGRESO'
toting = CDbl(Label14) + CDbl(Label15) + CDbl(Label29)
Label26 = Format(toting, "currency")

'TOTAL EGRESO'
toteg = CDbl(Label16) + CDbl(Label18) + CDbl(Label19) + CDbl(Label20) + CDbl(Label28)
Label27 = Format(toteg, "currency")

'TOTAL GRAL'
tot = CDbl(Label26) - CDbl(Label27)
TotalGral = Format(tot, "currency")
End Sub

Private Sub Command2_Click()
DTPicker2 = Date
Frame1.BorderStyle = 0

Call ALTA_INGRESO
Call ALTA_EGRESO

''' INGRESOS'''

With MSFlexGrid1

For i = 1 To .Rows - 1
    If .TextMatrix(i, 2) <> "EFECTIVO" Then
    If .TextMatrix(i, 2) = "" Or .TextMatrix(i, 3) = "" Then
        .TextMatrix(i, 2) = Format(0, "currency")
        .TextMatrix(i, 3) = Format(0, "currency")
        .TextMatrix(i, 4) = Format(0, "currency")
    Else
        efectivo_ing = CDbl(efectivo_ing) + .TextMatrix(i, 2)
        cheque_ing = CDbl(cheque_ing) + .TextMatrix(i, 3)
        insu_ing = CDbl(insu_ing) + .TextMatrix(i, 4)
    End If
    End If
Next
Label14 = Format(efectivo_ing, "currency")
Label15 = Format(cheque_ing, "currency")
Label29 = Format(insu_ing, "currency")
'Label22 = Label15
End With

'''EGRESOS'''

With MSFlexGrid2

For i = 1 To .Rows - 1
    If .TextMatrix(i, 2) <> "EFECTIVO" Then
    If .TextMatrix(i, 2) = "" Or .TextMatrix(i, 3) = "" Then
        .TextMatrix(i, 2) = Format(0, "currency")
        .TextMatrix(i, 3) = Format(0, "currency")
        .TextMatrix(i, 4) = Format(0, "currency")
        .TextMatrix(i, 5) = Format(0, "currency")
        .TextMatrix(i, 6) = Format(0, "currency")
    Else
        efectivo_sale = CDbl(efectivo_sale) + .TextMatrix(i, 2)
        cheque_pro = CDbl(cheque_pro) + .TextMatrix(i, 3)
        cheque_ter = CDbl(cheque_ter) + .TextMatrix(i, 4)
        reten = CDbl(reten) + .TextMatrix(i, 5)
        insumos = CDbl(insumos) + .TextMatrix(i, 6)
    End If
    End If
Next

End With

Label16 = Format(efectivo_sale, "currency")
Label18 = Format(cheque_pro, "currency")
Label28 = Format(cheque_ter, "currency")
Label19 = Format(insumos, "currency")
Label20 = Format(reten, "currency")
'Label23 = Label18
'Label24 = Label19
'Label25 = Label20
'subefect = CDbl(Label14) - CDbl(Label16)
'Label21 = Format(subefect, "currency")

'TOTAL INGRESO'
toting = CDbl(Label14) + CDbl(Label15) + CDbl(Label29)
Label26 = Format(toting, "currency")

'TOTAL EGRESO'
toteg = CDbl(Label16) + CDbl(Label18) + CDbl(Label19) + CDbl(Label20) + CDbl(Label28)
Label27 = Format(toteg, "currency")
'TOTAL GRAL'

tot = CDbl(Label26) - CDbl(Label27)
TotalGral = Format(tot, "currency")
End Sub

Private Sub Command3_Click()

Printer.Orientation = 1
Printer.FontSize = 8
Printer.Font = arial
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(130); "Fecha: "; Date
Printer.FontSize = 10
Printer.Print Tab(50); SISTEMA;
Printer.Print Tab(40); SISTEMA_DIR;
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "LISTADO DE INGRESOS Y EGRESOS"
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.Print Tab(10); Label5; "  "; DTPicker1; "  "; Label6; "  "; DTPicker2
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.FontSize = 9

Printer.Print Tab(27); Label1; Tab(45); Label2; Tab(75); Label3; Tab(105); Label7; Tab(130); Label8; ' Tab(125); "FECHA VTO"
Printer.Print Tab(10); ""
Printer.Print Tab(10); Label9; Tab(27); Label14; Tab(47); Label15; Tab(77); Tab(108); Label29; 'Tab(130); Label20; ' Tab(125); "FECHA VTO"
Printer.Print Tab(10); ""
Printer.Print Tab(10); Label10; Tab(27); Label16; Tab(47); Label28; Tab(77); Label18; Tab(108); Label19; Tab(130); Label20; ' Tab(125); "FECHA VTO"
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); Label12; "    "; Label26; Tab(77); Label17; "    "; Label27
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); Label13; "    "; TotalGral

Printer.EndDoc
End Sub

Private Sub Form_Load()

DTPicker2 = Date
Frame1.BorderStyle = 0

Call ALTA_INGRESO
Call ALTA_EGRESO

''' INGRESOS'''

With MSFlexGrid1

For i = 1 To .Rows - 1
    If .TextMatrix(i, 2) <> "EFECTIVO" Then
    If .TextMatrix(i, 2) = "" Or .TextMatrix(i, 3) = "" Then
        .TextMatrix(i, 2) = Format(0, "currency")
        .TextMatrix(i, 3) = Format(0, "currency")
        .TextMatrix(i, 4) = Format(0, "currency")
    Else
        efectivo_ing = CDbl(efectivo_ing) + .TextMatrix(i, 2)
        cheque_ing = CDbl(cheque_ing) + .TextMatrix(i, 3)
        insu_ing = CDbl(insu_ing) + .TextMatrix(i, 4)
        total = CDbl(.TextMatrix(i, 2)) + CDbl(.TextMatrix(i, 3)) + CDbl(.TextMatrix(i, 4))
        .TextMatrix(i, 1) = Format(total, "currency")
    End If
    End If
Next
Label14 = Format(efectivo_ing, "currency")
Label15 = Format(cheque_ing, "currency")
Label29 = Format(insu_ing, "currency")
'Label22 = Label15
End With

'''EGRESOS'''

With MSFlexGrid2

For i = 1 To .Rows - 1
    If .TextMatrix(i, 2) <> "EFECTIVO" Then
    If .TextMatrix(i, 2) = "" Or .TextMatrix(i, 3) = "" Then
        .TextMatrix(i, 2) = Format(0, "currency")
        .TextMatrix(i, 3) = Format(0, "currency")
        .TextMatrix(i, 4) = Format(0, "currency")
        .TextMatrix(i, 5) = Format(0, "currency")
        .TextMatrix(i, 6) = Format(0, "currency")
    Else
        efectivo_sale = CDbl(efectivo_sale) + .TextMatrix(i, 2)
        cheque_pro = CDbl(cheque_pro) + .TextMatrix(i, 3)
        cheque_ter = CDbl(cheque_ter) + .TextMatrix(i, 4)
        reten = CDbl(reten) + .TextMatrix(i, 5)
        insumos = CDbl(insumos) + .TextMatrix(i, 6)
        total = CDbl(.TextMatrix(i, 2)) + CDbl(.TextMatrix(i, 3)) + CDbl(.TextMatrix(i, 4)) + CDbl(.TextMatrix(i, 5)) + CDbl(.TextMatrix(i, 6))
        .TextMatrix(i, 1) = Format(total, "currency")
    End If
    End If
Next

End With

Label16 = Format(efectivo_sale, "currency")
Label18 = Format(cheque_pro, "currency")
Label28 = Format(cheque_ter, "currency")
Label19 = Format(insumos, "currency")
Label20 = Format(reten, "currency")
'Label23 = Label18
'Label24 = Label19
'Label25 = Label20
'subefect = CDbl(Label14) - CDbl(Label16)
'Label21 = Format(subefect, "currency")

'TOTAL INGRESO'
toting = CDbl(Label14) + CDbl(Label15) + CDbl(Label29)
Label26 = Format(toting, "currency")

'TOTAL EGRESO'
toteg = CDbl(Label16) + CDbl(Label18) + CDbl(Label19) + CDbl(Label20) + CDbl(Label28)
Label27 = Format(toteg, "currency")

'TOTAL GRAL'
tot = CDbl(Label26) - CDbl(Label27)
TotalGral = Format(tot, "currency")
End Sub
Private Sub INGRESOS()
With MSFlexGrid1
.FixedCols = 0
.Cols = 5
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "NOMBRE/RAZON SOCIAL"
.TextMatrix(0, 1) = "MONTO"
.TextMatrix(0, 2) = "0"
.TextMatrix(0, 3) = "0"
.TextMatrix(0, 4) = "0"


.ColWidth(0) = 3000
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 1500
'.ColWidth(4) = 2000
End With
End Sub
Private Sub EGRESOS()
With MSFlexGrid2
.FixedCols = 0
.Cols = 7
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "NOMBRE/ RAZON SOCIAL"
.TextMatrix(0, 1) = "MONTO"
.TextMatrix(0, 2) = "EFECTIVO"
.TextMatrix(0, 3) = "CHEQUES PROPIOS"
.TextMatrix(0, 4) = "CHEQUES TERCEROS"
.TextMatrix(0, 5) = "RETENCION"
.TextMatrix(0, 6) = "INSUMOS / TARJETAS"

.ColWidth(0) = 3000
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.ColWidth(6) = 1500
End With
End Sub
Private Sub ALTA_INGRESO()
'For i = 0 To MSFlexGrid2.Rows - 1
    SQL = "select *  from mov_clientes order by cliente"
    TABLA.Open SQL, conexion_BD
    Call INGRESOS
    With MSFlexGrid1
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!cliente

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
        
            cta2 = .TextMatrix(lin - 1, 2)
            cta3 = .TextMatrix(lin - 1, 3)
            cta4 = .TextMatrix(lin - 1, 4)
            
            Select Case TABLA!tipo
           
                Case "EFECTIVO"
                    'cta2 = 0
                    .TextMatrix(lin, 2) = TABLA!pago
                    cta2 = CDbl(.TextMatrix(lin - 1, 2)) + CDbl(.TextMatrix(lin, 2))
                    
                Case "CHEQUE"
                    'cta3 = 0
                    If TABLA!pago < 0 Then
                        cta3 = TABLA!pago * -1
                    Else
                        cta3 = TABLA!pago
                    End If
                    .TextMatrix(lin, 3) = cta3
                    cta3 = CDbl(.TextMatrix(lin, 3)) + CDbl(.TextMatrix(lin - 1, 3))
                
                Case Else
                    'cta4 = 0
                    .TextMatrix(lin, 4) = TABLA!pago
                    cta4 = CDbl(.TextMatrix(lin, 4)) + CDbl(.TextMatrix(lin - 1, 4))
                End Select
            
            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 2) = Format(cta2, "currency")
            .TextMatrix(lin, 3) = Format(cta3, "currency")
            .TextMatrix(lin, 4) = Format(cta4, "currency")
            'cta2 = 0
            'cta3 = 0
            'cta4 = 0
                        
        Else
 
            .TextMatrix(lin, 0) = TABLA!cliente
            
                Select Case TABLA!tipo
                
                    Case "EFECTIVO"
                        efec1 = TABLA!pago
                        .TextMatrix(lin, 2) = Format(efec1, "currency")
                        .TextMatrix(lin, 3) = Format(0, "currency")
                        .TextMatrix(lin, 4) = Format(0, "currency")
                
                    Case "CHEQUE"
                        If TABLA!pago < 0 Then
                            che1 = TABLA!pago * -1
                        Else
                            che1 = TABLA!pago
                        End If
                        .TextMatrix(lin, 2) = Format(0, "currency")
                        .TextMatrix(lin, 3) = Format(che1, "currency")
                        .TextMatrix(lin, 4) = Format(0, "currency")
                    
                    Case Else
                        cta4 = TABLA!pago
                        .TextMatrix(lin, 2) = Format(0, "currency")
                        .TextMatrix(lin, 3) = Format(0, "currency")
                        .TextMatrix(lin, 4) = Format(cta4, "currency")
                
                    End Select

                cta2 = 0
                cta3 = 0
                cta4 = 0
                efec1 = 0
                che1 = 0
        End If

        TABLA.MoveNext
    Loop
    TABLA.Close
.TextMatrix(0, 2) = "EFECTIVO"
.TextMatrix(0, 3) = "CHEQUE TERCEROS"
.TextMatrix(0, 4) = "INSUMOS"
   End With
   

End Sub
Private Sub ALTA_EGRESO()
SQL = "select * from mov_proveedor order by proveedor"
TABLA.Open SQL, conexion_BD

Call EGRESOS

Do While Not TABLA.EOF
    With MSFlexGrid2
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!proveedor
    
    If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
            
            efect = .TextMatrix(lin - 1, 2)
            cheCH = .TextMatrix(lin - 1, 3)
            cheVT = .TextMatrix(lin - 1, 4)
            insu = .TextMatrix(lin - 1, 6)
            
        Select Case TABLA!tipo
            Case "EFECTIVO"
                efect = 0
                .TextMatrix(lin, 2) = TABLA!pago
                efect = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
           
            Case "CHEQUE CH"
                cheCH = 0
                .TextMatrix(lin, 3) = TABLA!pago
                cheCH = CDbl(.TextMatrix(lin, 3)) + CDbl(.TextMatrix(lin - 1, 3))
            
            Case "CHEQUE VT"
                cheVT = 0
                .TextMatrix(lin, 4) = TABLA!pago
                cheVT = CDbl(.TextMatrix(lin, 4)) + CDbl(.TextMatrix(lin - 1, 4))
            
            'Case "RETENCION"
            '    ret = "RE"
            '    .TextMatrix(lin, 5) = "RETEN"
            '    .TextMatrix(lin, 5) = 0
            
            Case Else
                insu = 0
                .TextMatrix(lin, 6) = TABLA!pago
                insu = CDbl(.TextMatrix(lin, 6)) + CDbl(.TextMatrix(lin - 1, 6))
        
        End Select
        
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        .TextMatrix(lin, 2) = Format(efect, "currency")
        .TextMatrix(lin, 3) = Format(cheCH, "currency")
        .TextMatrix(lin, 4) = Format(cheVT, "currency")
        .TextMatrix(lin, 5) = Format(0, "currency")
        .TextMatrix(lin, 6) = Format(insu, "currency")

    
    Else
        Select Case TABLA!tipo
            Case "EFECTIVO"
                efect = TABLA!pago
                .TextMatrix(lin, 1) = Format(0, "currency")
                .TextMatrix(lin, 2) = Format(efect, "currency")
                .TextMatrix(lin, 3) = Format(0, "currency")
                .TextMatrix(lin, 4) = Format(0, "currency")
                .TextMatrix(lin, 5) = Format(0, "currency")
                .TextMatrix(lin, 6) = Format(0, "currency")
            
            Case "CHEQUE CH"
                cheCH = TABLA!pago
                .TextMatrix(lin, 1) = Format(0, "currency")
                .TextMatrix(lin, 2) = Format(0, "currency")
                .TextMatrix(lin, 3) = Format(cheCH, "currency")
                .TextMatrix(lin, 4) = Format(0, "currency")
                .TextMatrix(lin, 5) = Format(0, "currency")
                .TextMatrix(lin, 6) = Format(0, "currency")
                
            Case "CHEQUE VT"
                cheVT = TABLA!pago
                .TextMatrix(lin, 1) = Format(0, "currency")
                .TextMatrix(lin, 2) = Format(0, "currency")
                .TextMatrix(lin, 3) = Format(0, "currency")
                .TextMatrix(lin, 4) = Format(cheVT, "currency")
                .TextMatrix(lin, 5) = Format(0, "currency")
                .TextMatrix(lin, 6) = Format(0, "currency")

            'Case "RETENCION"
            '    .TextMatrix(lin, 5) = "retencion"
            
            Case Else
                insu = TABLA!pago
                .TextMatrix(lin, 1) = Format(0, "currency")
                .TextMatrix(lin, 2) = Format(0, "currency")
                .TextMatrix(lin, 3) = Format(0, "currency")
                .TextMatrix(lin, 4) = Format(0, "currency")
                .TextMatrix(lin, 4) = Format(0, "currency")
                .TextMatrix(lin, 5) = Format(0, "currency")
                .TextMatrix(lin, 6) = Format(insu, "currency")
        
        End Select
        efect = 0
        cheVT = 0
        cheCH = 0
        insu = 0
    End If
    TABLA.MoveNext
    End With
Loop
    
TABLA.Close

''' RETENCIONES'''

For i = 0 To MSFlexGrid2.Rows - 1
    SQL = "select *  from retencion order by destino"
    TABLA.Open SQL, conexion_BD

    Do While Not TABLA.EOF
  
        If MSFlexGrid2.TextMatrix(i, 0) = TABLA!destino Then
            ret = CDbl(ret) + TABLA!importe
            MSFlexGrid2.TextMatrix(i, 5) = Format(ret, "currency")
        Else
            ret = Format(0, "currency")

        End If
        TABLA.MoveNext
    Loop
TABLA.Close
Next


            
            
SQL = "select * from mov_rrhh order by nombre_rrhh"
TABLA.Open SQL, conexion_BD

Do While Not TABLA.EOF
    With MSFlexGrid2
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!nombre_rrhh
    
    If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
                
        .TextMatrix(lin, 2) = TABLA!adelanto
        adelan = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
        
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        .TextMatrix(lin, 2) = Format(adelan, "currency")
        .TextMatrix(lin, 3) = Format(0, "currency")
        .TextMatrix(lin, 4) = Format(0, "currency")
        .TextMatrix(lin, 5) = Format(0, "currency")
        .TextMatrix(lin, 6) = Format(0, "currency")
    
    Else
        adelan = TABLA!adelanto
        .TextMatrix(lin, 2) = Format(adelan, "currency")
        .TextMatrix(lin, 3) = Format(0, "currency")
        .TextMatrix(lin, 4) = Format(0, "currency")
        .TextMatrix(lin, 5) = Format(0, "currency")
        .TextMatrix(lin, 6) = Format(0, "currency")
    End If
    End With
    TABLA.MoveNext
Loop
TABLA.Close
End Sub
Private Sub ALTA_INGRESO_FECHA()

Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)


SQL = "select * from mov_clientes where fecha >= " + CStr(inicio) + " and fecha <=  " + CStr(final) + " order by cliente"
TABLA.Open SQL, conexion_BD
    
    Call INGRESOS
    With MSFlexGrid1
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!cliente

        'SI EL NOMBRE DEL REG NUEVO ES = AL ANTERIOR SUMA EL MONTO Y BORRA EL REG NUEVO
    
        If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
            
            cta2 = .TextMatrix(lin - 1, 2)
            cta3 = .TextMatrix(lin - 1, 3)
            cta4 = .TextMatrix(lin - 1, 4)
            
            Select Case TABLA!tipo
                            
                Case "EFECTIVO"
                    cta2 = 0
                    .TextMatrix(lin, 2) = TABLA!pago
                    cta2 = CDbl(.TextMatrix(lin - 1, 2)) + CDbl(.TextMatrix(lin, 2))
                    
                Case "CHEQUE"
                    cta3 = 0
                    If TABLA!pago < 0 Then
                        cta3 = TABLA!pago * -1
                    Else
                        cta3 = TABLA!pago
                    End If
                    .TextMatrix(lin, 3) = cta3
                    cta3 = CDbl(.TextMatrix(lin, 3)) + CDbl(.TextMatrix(lin - 1, 3))
                
                Case Else
                    cta4 = 0
                    .TextMatrix(lin, 4) = TABLA!pago
                    cta4 = CDbl(.TextMatrix(lin, 3)) + CDbl(.TextMatrix(lin - 1, 3))
                End Select

            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 2) = Format(cta2, "currency")
            .TextMatrix(lin, 3) = Format(cta3, "currency")
            .TextMatrix(lin, 4) = Format(cta4, "currency")
            
        Else
 
            .TextMatrix(lin, 0) = TABLA!cliente
                
                Select Case TABLA!tipo
                
                    Case "EFECTIVO"
                        efect = TABLA!pago
                        .TextMatrix(lin, 2) = Format(efect, "currency")
                        .TextMatrix(lin, 3) = Format(0, "currency")
                        .TextMatrix(lin, 4) = Format(0, "currency")
                
                    Case "CHEQUE"
                        If TABLA!pago < 0 Then
                            cheque = TABLA!pago * -1
                        Else
                            cheque = TABLA!pago
                        End If
                        .TextMatrix(lin, 2) = Format(0, "currency")
                        .TextMatrix(lin, 3) = Format(cheque, "currency")
                        .TextMatrix(lin, 4) = Format(0, "currency")
                    
                    Case Else
                        nose = TABLA!pago
                        .TextMatrix(lin, 2) = Format(0, "currency")
                        .TextMatrix(lin, 3) = Format(0, "currency")
                        .TextMatrix(lin, 4) = Format(nose, "currency")
                
                    End Select

                cta2 = 0
                cta3 = 0
                cta4 = 0
                
        End If

        TABLA.MoveNext
    Loop
    TABLA.Close
    
.TextMatrix(0, 2) = "EFECTIVO"
.TextMatrix(0, 3) = "CHEQUE TERCEROS"
.TextMatrix(0, 4) = "INSUMOS"

End With
End Sub
Private Sub ALTA_EGRESO_FECHA()

Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)


SQL = "select * from mov_proveedor where fecha >= " + CStr(inicio) + " and fecha <=  " + CStr(final) + " order by proveedor"
TABLA.Open SQL, conexion_BD

Call EGRESOS

Do While Not TABLA.EOF
    With MSFlexGrid2
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!proveedor
    
    If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
    
        efect = .TextMatrix(lin - 1, 2)
        cheCH = .TextMatrix(lin - 1, 3)
        cheVT = .TextMatrix(lin - 1, 4)
        insu = .TextMatrix(lin - 1, 6)
            
        Select Case TABLA!tipo
            Case "EFECTIVO"
                efect = 0
                .TextMatrix(lin, 2) = TABLA!pago
                efect = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
            
            Case "CHEQUE CH"
                cheCH = 0
                .TextMatrix(lin, 3) = TABLA!pago
                cheCH = CDbl(.TextMatrix(lin, 3)) + CDbl(.TextMatrix(lin - 1, 3))
                
            Case "CHEQUE VT"
                cheVT = 0
                .TextMatrix(lin, 4) = TABLA!pago
                cheVT = CDbl(.TextMatrix(lin, 4)) + CDbl(.TextMatrix(lin - 1, 4))
                
            'Case "RETENCION"
            '    ret = "RE"
            '    .TextMatrix(lin, 5) = "RETEN"
            '    .TextMatrix(lin, 6) = 0
            
            Case Else
                insu = 0
                .TextMatrix(lin, 6) = TABLA!pago
                insu = CDbl(.TextMatrix(lin, 6)) + CDbl(.TextMatrix(lin - 1, 6))
        
        End Select
        
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        .TextMatrix(lin, 2) = Format(efect, "currency")
        .TextMatrix(lin, 3) = Format(cheCH, "currency")
        .TextMatrix(lin, 4) = Format(cheVT, "currency")
        .TextMatrix(lin, 5) = Format(0, "currency")
        .TextMatrix(lin, 6) = Format(insu, "currency")

    
    Else
        Select Case TABLA!tipo
            Case "EFECTIVO"
                efect = TABLA!pago
                .TextMatrix(lin, 1) = Format(0, "currency")
                .TextMatrix(lin, 2) = Format(efect, "currency")
                .TextMatrix(lin, 3) = Format(0, "currency")
                .TextMatrix(lin, 4) = Format(0, "currency")
                .TextMatrix(lin, 5) = Format(0, "currency")
                .TextMatrix(lin, 6) = Format(0, "currency")
            
            Case "CHEQUE CH"
                cheCH = TABLA!pago
                .TextMatrix(lin, 1) = Format(0, "currency")
                .TextMatrix(lin, 2) = Format(0, "currency")
                .TextMatrix(lin, 3) = Format(cheCH, "currency")
                .TextMatrix(lin, 4) = Format(0, "currency")
                .TextMatrix(lin, 5) = Format(0, "currency")
                .TextMatrix(lin, 6) = Format(0, "currency")
                
            Case "CHEQUE VT"
                cheVT = TABLA!pago
                .TextMatrix(lin, 1) = Format(0, "currency")
                .TextMatrix(lin, 2) = Format(0, "currency")
                .TextMatrix(lin, 3) = Format(0, "currency")
                .TextMatrix(lin, 4) = Format(cheVT, "currency")
                .TextMatrix(lin, 5) = Format(0, "currency")
                .TextMatrix(lin, 6) = Format(0, "currency")

            'Case "RETENCION"

            '    .TextMatrix(lin, 5) = "retencion"
            
            Case Else
                nose = TABLA!pago
                .TextMatrix(lin, 1) = Format(0, "currency")
                .TextMatrix(lin, 2) = Format(0, "currency")
                .TextMatrix(lin, 3) = Format(0, "currency")
                .TextMatrix(lin, 4) = Format(0, "currency")
                .TextMatrix(lin, 5) = Format(0, "currency")
                .TextMatrix(lin, 6) = Format(nose, "currency")
        
        End Select
        efect = 0
        cheCH = 0
        cheVT = 0
        insu = 0
    End If
    TABLA.MoveNext
    End With
Loop
    
TABLA.Close

''' RETENCIONES'''

For i = 0 To MSFlexGrid2.Rows - 1
    SQL = "select * from retencion where fecha >= " + CStr(inicio) + " and fecha <=  " + CStr(final) + " order by destino"
    TABLA.Open SQL, conexion_BD

    Do While Not TABLA.EOF
  
        If MSFlexGrid2.TextMatrix(i, 0) = TABLA!destino Then
            ret = TABLA!importe
            MSFlexGrid2.TextMatrix(i, 5) = Format(ret, "currency")
        End If
        TABLA.MoveNext
    Loop
TABLA.Close
Next



SQL = "select * from mov_rrhh where fecha_mod >= " + CStr(inicio) + " and fecha_mod <=  " + CStr(final) + " order by fecha_mod"
TABLA.Open SQL, conexion_BD

Do While Not TABLA.EOF
    With MSFlexGrid2
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!nombre_rrhh
    
    If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
                
        .TextMatrix(lin, 2) = TABLA!adelanto
        adelan = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
        
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        .TextMatrix(lin, 2) = Format(adelan, "currency")
        .TextMatrix(lin, 3) = Format(0, "currency")
        .TextMatrix(lin, 4) = Format(0, "currency")
        .TextMatrix(lin, 5) = Format(0, "currency")
        .TextMatrix(lin, 6) = Format(0, "currency")
    
    Else
        adelan = TABLA!adelanto
        .TextMatrix(lin, 2) = Format(adelan, "currency")
        .TextMatrix(lin, 3) = Format(0, "currency")
        .TextMatrix(lin, 4) = Format(0, "currency")
        .TextMatrix(lin, 5) = Format(0, "currency")
        .TextMatrix(lin, 6) = Format(0, "currency")
    End If
    End With
    TABLA.MoveNext
Loop
TABLA.Close
End Sub


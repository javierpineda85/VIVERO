VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Listado_cheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historico de Cheques"
   ClientHeight    =   9360
   ClientLeft      =   765
   ClientTop       =   870
   ClientWidth     =   13470
   Icon            =   "Listado_cheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13470
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   3960
      TabIndex        =   25
      Top             =   1680
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton Command7 
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
         Left            =   6720
         Picture         =   "Listado_cheques.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Listado_cheques.frx":0B14
         DragIcon        =   "Listado_cheques.frx":109E
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
         Picture         =   "Listado_cheques.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         TabIndex        =   29
         Top             =   315
         Width           =   1935
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
         Height          =   330
         Left            =   840
         TabIndex        =   27
         Top             =   315
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "hasta:"
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
         Left            =   2880
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Desde"
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
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command5 
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
      Left            =   11760
      Picture         =   "Listado_cheques.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   4200
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Listado_cheques.frx":213C
         DragIcon        =   "Listado_cheques.frx":26C6
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
         Picture         =   "Listado_cheques.frx":2C50
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
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
         Left            =   6480
         Picture         =   "Listado_cheques.frx":31DA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
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
         Left            =   2640
         TabIndex        =   7
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Buscar por Nº de cheque:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   16
      Top             =   1680
      Width           =   3375
      Begin VB.OptionButton Option3 
         Caption         =   "Rango de Números"
         Height          =   495
         Left            =   2160
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Número"
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
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fecha"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   4200
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton Command3 
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
         Left            =   6480
         Picture         =   "Listado_cheques.frx":3764
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Listado_cheques.frx":3CEE
         DragIcon        =   "Listado_cheques.frx":4278
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
         Picture         =   "Listado_cheques.frx":4802
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3840
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
         Format          =   61472769
         CurrentDate     =   41087
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
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
         Format          =   61472769
         CurrentDate     =   40909
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta:"
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
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Buscar desde:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5775
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   10186
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
   Begin VB.Label Label10 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   8640
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Cheques Rechazados"
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
      Left            =   4560
      TabIndex        =   22
      Top             =   8640
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Cheques Ingresados al Banco"
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
      Left            =   1080
      TabIndex        =   20
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   8640
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      Top             =   8640
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO HISTORICO DE CHEQUES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4080
      TabIndex        =   10
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "Listado_cheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
SQL = " SELECT entracheque.n_interno,entracheque.n_cheque,entracheque.importe,entracheque.banco,entracheque.fecha_depo,cliente,entracheque.fecha_vto,mov_cheques.destino,entracheque.rechazado,entracheque.cuit FROM entracheque LEFT JOIN mov_cheques ON entracheque.n_cheque=mov_cheques.n_cheque ORDER BY entracheque.fecha_vto"
TABLA.Open SQL, conexion_BD


Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
    If IsNull(TABLA!destino) Then
        MSFlexGrid1.TextMatrix(lin, 7) = "NO CONSIGNADO"
    Else
        MSFlexGrid1.TextMatrix(lin, 7) = TABLA!destino
    End If
    
    If IsNull(TABLA!cuit) Then
        MSFlexGrid1.TextMatrix(lin, 8) = "-"
    Else
        MSFlexGrid1.TextMatrix(lin, 8) = TABLA!cuit
    End If
    
    If TABLA!importe < "0" Then
        With MSFlexGrid1
        num = CDbl(.TextMatrix(lin, 2)) * -1
        .TextMatrix(lin, 2) = Format(num, "currency")
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 1
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 2
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 3
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 4
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 5
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 6
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 7
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 8
        .CellBackColor = &HFF00&
        End With
    End If
    
    If TABLA!rechazado = "-1" Then
        With MSFlexGrid1
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF&
        .Col = 1
        .CellBackColor = &HFF&
        .Col = 2
        .CellBackColor = &HFF&
        .Col = 3
        .CellBackColor = &HFF&
        .Col = 4
        .CellBackColor = &HFF&
        .Col = 5
        .CellBackColor = &HFF&
        .Col = 6
        .CellBackColor = &HFF&
        .Col = 7
        .CellBackColor = &HFF&
        .Col = 8
        .CellBackColor = &HFF&
        End With
    End If
    If TABLA!rechazado = 0 Then
    
        importe = CDbl(importe) + CDbl(MSFlexGrid1.TextMatrix(lin, 2))
    End If
    TABLA.MoveNext
Loop
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub Command2_Click()
Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)


SQL = " SELECT entracheque.n_interno,entracheque.n_cheque,entracheque.importe,entracheque.banco,entracheque.fecha_depo,cliente,entracheque.fecha_vto,mov_cheques.destino,entracheque.rechazado, entracheque.cuit FROM entracheque LEFT JOIN mov_cheques ON entracheque.n_cheque=mov_cheques.n_cheque where entracheque.fecha_vto >= " + CStr(inicio) + " and entracheque.fecha_vto <= " + CStr(final) + " ORDER BY entracheque.fecha_vto"

TABLA.Open SQL, conexion_BD


MSFlexGrid1.Clear

Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
    If IsNull(TABLA!destino) Then
        MSFlexGrid1.TextMatrix(lin, 7) = "NO CONSIGNADO"
    Else
        MSFlexGrid1.TextMatrix(lin, 7) = TABLA!destino
    End If
    
    If IsNull(TABLA!cuit) Then
        MSFlexGrid1.TextMatrix(lin, 8) = "-"
    Else
        MSFlexGrid1.TextMatrix(lin, 8) = TABLA!cuit
    End If
    
    If TABLA!importe < "0" Then
        With MSFlexGrid1
        num = CDbl(.TextMatrix(lin, 2)) * -1
        .TextMatrix(lin, 2) = Format(num, "currency")
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 1
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 2
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 3
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 4
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 5
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 6
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 7
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 8
        .CellBackColor = &HFF00&
        End With
    End If
    
    If TABLA!rechazado = "-1" Then
        With MSFlexGrid1
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF&
        .Col = 1
        .CellBackColor = &HFF&
        .Col = 2
        .CellBackColor = &HFF&
        .Col = 3
        .CellBackColor = &HFF&
        .Col = 4
        .CellBackColor = &HFF&
        .Col = 5
        .CellBackColor = &HFF&
        .Col = 6
        .CellBackColor = &HFF&
        .Col = 7
        .CellBackColor = &HFF&
        .Col = 8
        .CellBackColor = &HFF&
        End With
    End If
    If TABLA!rechazado = 0 Then
    
        importe = CDbl(importe) + CDbl(MSFlexGrid1.TextMatrix(lin, 2))
    End If
    TABLA.MoveNext
Loop
'Label3 = importe
TABLA.Close

End Sub

Private Sub Command3_Click()
'P = "select * from entracheque order by fecha_vto"
'TABLA.Open P, conexion_BD

SQL = " SELECT entracheque.n_interno,entracheque.n_cheque,entracheque.importe,entracheque.banco,entracheque.fecha_depo,cliente,entracheque.fecha_vto,mov_cheques.destino,entracheque.rechazado,entracheque.cuit FROM entracheque LEFT JOIN mov_cheques ON entracheque.n_cheque=mov_cheques.n_cheque ORDER BY entracheque.fecha_vto"
TABLA.Open SQL, conexion_BD
Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
    If IsNull(TABLA!destino) Then
        MSFlexGrid1.TextMatrix(lin, 7) = "NO CONSIGNADO"
    Else
        MSFlexGrid1.TextMatrix(lin, 7) = TABLA!destino
    End If
    
    If IsNull(TABLA!cuit) Then
        MSFlexGrid1.TextMatrix(lin, 8) = "-"
    Else
        MSFlexGrid1.TextMatrix(lin, 8) = TABLA!cuit
    End If
    
    If TABLA!importe < "0" Then
        With MSFlexGrid1
        num = CDbl(.TextMatrix(lin, 2)) * -1
        .TextMatrix(lin, 2) = Format(num, "currency")
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 1
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 2
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 3
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 4
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 5
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 6
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 7
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 8
        .CellBackColor = &HFF00&
        End With
    End If
    
    If TABLA!rechazado = "-1" Then
        With MSFlexGrid1
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF&
        .Col = 1
        .CellBackColor = &HFF&
        .Col = 2
        .CellBackColor = &HFF&
        .Col = 3
        .CellBackColor = &HFF&
        .Col = 4
        .CellBackColor = &HFF&
        .Col = 5
        .CellBackColor = &HFF&
        .Col = 6
        .CellBackColor = &HFF&
        .Col = 7
        .CellBackColor = &HFF&
        .Col = 8
        .CellBackColor = &HFF&
        End With
    End If
    If TABLA!rechazado = 0 Then
    
        importe = CDbl(importe) + CDbl(MSFlexGrid1.TextMatrix(lin, 2))
    End If
    TABLA.MoveNext
Loop
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub Command4_Click()
'a = "select * from entracheque where n_cheque = " & Val(Text1) & " order by fecha_vto"
'TABLA.Open a, conexion_BD

SQL = " SELECT entracheque.n_interno,entracheque.n_cheque,entracheque.importe,entracheque.banco,entracheque.fecha_depo,cliente,entracheque.fecha_vto,mov_cheques.destino,entracheque.rechazado,entracheque.cuit FROM entracheque LEFT JOIN mov_cheques ON entracheque.n_cheque=mov_cheques.n_cheque WHERE mov_cheques.n_cheque=" & Val(Text1) & ""
TABLA.Open SQL, conexion_BD

MSFlexGrid1.Clear

Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
    If IsNull(TABLA!destino) Then
        MSFlexGrid1.TextMatrix(lin, 7) = "NO CONSIGNADO"
    Else
        MSFlexGrid1.TextMatrix(lin, 7) = TABLA!destino
    End If
    
    If IsNull(TABLA!cuit) Then
        MSFlexGrid1.TextMatrix(lin, 8) = "-"
    Else
        MSFlexGrid1.TextMatrix(lin, 8) = TABLA!cuit
    End If
    
    If TABLA!importe < "0" Then
        With MSFlexGrid1
        num = CDbl(.TextMatrix(lin, 2)) * -1
        .TextMatrix(lin, 2) = Format(num, "currency")
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 1
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 2
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 3
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 4
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 5
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 6
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 7
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 8
        .CellBackColor = &HFF00&
        End With
    End If
    If TABLA!rechazado = "-1" Then
        With MSFlexGrid1
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF&
        .Col = 1
        .CellBackColor = &HFF&
        .Col = 2
        .CellBackColor = &HFF&
        .Col = 3
        .CellBackColor = &HFF&
        .Col = 4
        .CellBackColor = &HFF&
        .Col = 5
        .CellBackColor = &HFF&
        .Col = 6
        .CellBackColor = &HFF&
        .Col = 7
        .CellBackColor = &HFF&
        .Col = 8
        .CellBackColor = &HFF&
        End With
    End If
    If TABLA!rechazado = "0" Then
    
        importe = CDbl(importe) + CDbl(MSFlexGrid1.TextMatrix(lin, 2))
    End If
    TABLA.MoveNext
Loop
'Label3 = importe
TABLA.Close
End Sub

Private Sub Command5_Click()
Printer.Orientation = 1
Printer.FontSize = 10
Printer.Font = arial
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(10); "Fecha: "; Date; Tab(57); SISTEMA; Tab(90); Page
Printer.Print Tab(50); SISTEMA_DIR
'Printer.Print Tab(10); "Fecha: "; Date; Tab(57); " STELLA DAVIRE"; Tab(90); Page
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(30); "LISTADO HISTORICOS DE CHEQUES"
Printer.Print Tab(10); ""
Printer.Print Tab(10); " Detalle desde "; DTPicker1; " Hasta "; DTPicker2
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.FontSize = 8
Printer.Print Tab(10); "Nº INTERNO"; Tab(25); "Nº CHEQUE"; Tab(43); "IMPORTE"; Tab(65); "CLIENTE"; Tab(108); "BANCO"; Tab(125); "FECHA VTO"
Printer.Print Tab(10); "======================================================================================================"


For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    interno = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    cheque = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    importe = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    cliente = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
    banco = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 4)
    vto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 5)

Printer.Print Tab(12); interno; Tab(27); cheque; Tab(45); importe; Tab(57); cliente; Tab(105); banco; Tab(127); vto; 'Tab(130); modif
Printer.Print Tab(10); "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

If i = 32 Or i = 68 Or i = 103 Or i = 138 Then
    Printer.NewPage
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); ""
    Printer.Print Tab(10); "Nº INTERNO"; Tab(25); "Nº CHEQUE"; Tab(43); "IMPORTE"; Tab(65); "CLIENTE"; Tab(108); "BANCO"; Tab(125); "FECHA VTO"
    Printer.Print Tab(10); "======================================================================================================"
End If

Next
Printer.Print Tab(10); ""
Printer.Print Tab(10); " TOTAL:"; Label3
Printer.EndDoc

End Sub

Private Sub Command6_Click()
Dim ppio As Long
Dim fin As Long

ppio = Text2
fin = Text3


SQL = " SELECT entracheque.n_interno,entracheque.n_cheque,entracheque.importe,entracheque.banco,entracheque.fecha_depo,cliente,entracheque.fecha_vto,mov_cheques.destino,entracheque.rechazado,entracheque.cuit FROM entracheque LEFT JOIN mov_cheques ON entracheque.n_cheque=mov_cheques.n_cheque where entracheque.n_cheque >= " + CStr(ppio) + " and entracheque.n_cheque <= " + CStr(fin) + " ORDER BY entracheque.n_cheque"

TABLA.Open SQL, conexion_BD


MSFlexGrid1.Clear

Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
    If IsNull(TABLA!destino) Then
        MSFlexGrid1.TextMatrix(lin, 7) = "NO CONSIGNADO"
    Else
        MSFlexGrid1.TextMatrix(lin, 7) = TABLA!destino
    End If
    
    If IsNull(TABLA!cuit) Then
        MSFlexGrid1.TextMatrix(lin, 8) = "-"
    Else
        MSFlexGrid1.TextMatrix(lin, 8) = TABLA!cuit
    End If

    If TABLA!importe < "0" Then
        With MSFlexGrid1
        num = CDbl(.TextMatrix(lin, 2)) * -1
        .TextMatrix(lin, 2) = Format(num, "currency")
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 1
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 2
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 3
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 4
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 5
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 6
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 7
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 8
        .CellBackColor = &HFF00&
        End With
    End If
    
    If TABLA!rechazado = "-1" Then
        With MSFlexGrid1
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF&
        .Col = 1
        .CellBackColor = &HFF&
        .Col = 2
        .CellBackColor = &HFF&
        .Col = 3
        .CellBackColor = &HFF&
        .Col = 4
        .CellBackColor = &HFF&
        .Col = 5
        .CellBackColor = &HFF&
        .Col = 6
        .CellBackColor = &HFF&
        .Col = 7
        .CellBackColor = &HFF&
        .Col = 8
        .CellBackColor = &HFF&
        End With
    End If
    If TABLA!rechazado = 0 Then
    
        importe = CDbl(importe) + CDbl(MSFlexGrid1.TextMatrix(lin, 2))
    End If
    TABLA.MoveNext
Loop
'Label3 = importe
TABLA.Close
End Sub

Private Sub Command7_Click()
SQL = " SELECT entracheque.n_interno,entracheque.n_cheque,entracheque.importe,entracheque.banco,entracheque.fecha_depo,cliente,entracheque.fecha_vto,mov_cheques.destino,entracheque.rechazado,entracheque.cuit FROM entracheque LEFT JOIN mov_cheques ON entracheque.n_cheque=mov_cheques.n_cheque ORDER BY entracheque.fecha_vto"
TABLA.Open SQL, conexion_BD


Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
    If IsNull(TABLA!destino) Then
        MSFlexGrid1.TextMatrix(lin, 7) = "NO CONSIGNADO"
    Else
        MSFlexGrid1.TextMatrix(lin, 7) = TABLA!destino
    End If
    
    If IsNull(TABLA!cuit) Then
        MSFlexGrid1.TextMatrix(lin, 8) = "-"
    Else
        MSFlexGrid1.TextMatrix(lin, 8) = TABLA!cuit
    End If
    
    If TABLA!importe < "0" Then
        With MSFlexGrid1
        num = CDbl(.TextMatrix(lin, 2)) * -1
        .TextMatrix(lin, 2) = Format(num, "currency")
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 1
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 2
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 3
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 4
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 5
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 6
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 7
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 8
        .CellBackColor = &HFF00&
        End With
    End If
    
    If TABLA!rechazado = "-1" Then
        With MSFlexGrid1
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF&
        .Col = 1
        .CellBackColor = &HFF&
        .Col = 2
        .CellBackColor = &HFF&
        .Col = 3
        .CellBackColor = &HFF&
        .Col = 4
        .CellBackColor = &HFF&
        .Col = 5
        .CellBackColor = &HFF&
        .Col = 6
        .CellBackColor = &HFF&
        .Col = 7
        .CellBackColor = &HFF&
        .Col = 8
        .CellBackColor = &HFF&
        End With
    End If
    If TABLA!rechazado = 0 Then
    
        importe = CDbl(importe) + CDbl(MSFlexGrid1.TextMatrix(lin, 2))
    End If
    TABLA.MoveNext
Loop
Label3 = Format(importe, "currency")
TABLA.Close
End Sub

Private Sub Form_Load()
DTPicker2 = Date

SQL = " SELECT entracheque.n_interno,entracheque.n_cheque,entracheque.importe,entracheque.banco,entracheque.fecha_depo,cliente,entracheque.fecha_vto,mov_cheques.destino,entracheque.rechazado,entracheque.cuit FROM entracheque LEFT JOIN mov_cheques ON entracheque.n_cheque=mov_cheques.n_cheque ORDER BY entracheque.fecha_vto"
TABLA.Open SQL, conexion_BD


Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
       
    If IsNull(TABLA!destino) Then
        MSFlexGrid1.TextMatrix(lin, 7) = "NO CONSIGNADO"
    Else
        MSFlexGrid1.TextMatrix(lin, 7) = TABLA!destino
    End If
    
    If IsNull(TABLA!cuit) Then
        MSFlexGrid1.TextMatrix(lin, 8) = "-"
    Else
        MSFlexGrid1.TextMatrix(lin, 8) = TABLA!cuit
    End If

    If TABLA!importe < "0" Then
        With MSFlexGrid1
        num = CDbl(.TextMatrix(lin, 2)) * -1
        .TextMatrix(lin, 2) = Format(num, "currency")
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 1
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 2
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 3
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 4
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 5
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 6
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 7
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 8
        .CellBackColor = &HFF00&
        End With
    End If
    
    If TABLA!rechazado = "-1" Then
        With MSFlexGrid1
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF&
        .Col = 1
        .CellBackColor = &HFF&
        .Col = 2
        .CellBackColor = &HFF&
        .Col = 3
        .CellBackColor = &HFF&
        .Col = 4
        .CellBackColor = &HFF&
        .Col = 5
        .CellBackColor = &HFF&
        .Col = 6
        .CellBackColor = &HFF&
        .Col = 7
        .CellBackColor = &HFF&
        .Col = 8
        .CellBackColor = &HFF&

        End With
    End If
    If TABLA!rechazado = 0 Then
    
        importe = CDbl(importe) + CDbl(MSFlexGrid1.TextMatrix(lin, 2))
    End If
    
    TABLA.MoveNext
Loop
Label3 = Format(importe, "currency")
TABLA.Close

End Sub
Private Sub Alta_Cheques()

With MSFlexGrid1
.FixedCols = 0
.Cols = 9
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "Nº INTERNO"
.TextMatrix(0, 1) = "Nº DE CHEQUE"
.TextMatrix(0, 2) = "IMPORTE"
.TextMatrix(0, 3) = "CLIENTE"
.TextMatrix(0, 4) = "BANCO"
.TextMatrix(0, 5) = "FECHA DE VTO"
.TextMatrix(0, 6) = "FECHA TRANSACCION"
.TextMatrix(0, 7) = "DESTINO"
.TextMatrix(0, 8) = "C.U.I.T"

.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 3000
.ColWidth(4) = 2000
.ColWidth(5) = 1900
.ColWidth(6) = 1900
.ColWidth(7) = 3000
.ColWidth(8) = 1900
End With
End Sub

Private Sub Option1_Click()
If Option1 = True Then
    Frame2.Visible = True
    Frame2.BorderStyle = 0
    Frame3.Visible = False
    Frame4.Visible = False
Else
    
    Frame3.Visible = True
    Frame3.BorderStyle = 0
    Frame2.Visible = False
    Frame4.Visible = False
    
End If

End Sub

Private Sub Option2_Click()
If Option2 = True Then
    Frame3.Visible = True
    Frame3.BorderStyle = 0
    Frame2.Visible = False
    Frame4.Visible = False
Else
    
    Frame2.Visible = True
    Frame2.BorderStyle = 0
    Frame3.Visible = False
    Frame4.Visible = False
    
End If
End Sub

Private Sub Option3_Click()
If Option3 = True Then
    Frame4.BorderStyle = 0
    Frame4.Visible = True
    Frame3.Visible = False
    Frame2.Visible = False
Else
    Frame4.Visible = False
    Frame3.Visible = False
    Frame2.Visible = False
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

SQL = " SELECT entracheque.n_interno,entracheque.n_cheque,entracheque.importe,entracheque.banco,entracheque.fecha_depo,cliente,entracheque.fecha_vto,mov_cheques.destino,entracheque.rechazado,entracheque.cuit FROM entracheque LEFT JOIN mov_cheques ON entracheque.n_cheque=mov_cheques.n_cheque WHERE mov_cheques.n_cheque=" & Val(Text1) & ""
TABLA.Open SQL, conexion_BD

MSFlexGrid1.Clear

Call Alta_Cheques

Do While Not TABLA.EOF
    lin = lin + 1
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(lin, 0) = TABLA!n_interno
    MSFlexGrid1.TextMatrix(lin, 1) = TABLA!n_cheque
    MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!importe, "currency")
    MSFlexGrid1.TextMatrix(lin, 3) = TABLA!cliente
    MSFlexGrid1.TextMatrix(lin, 4) = TABLA!banco
    MSFlexGrid1.TextMatrix(lin, 5) = TABLA!fecha_vto
    MSFlexGrid1.TextMatrix(lin, 6) = TABLA!fecha_depo
    If IsNull(TABLA!destino) Then
        MSFlexGrid1.TextMatrix(lin, 7) = "NO CONSIGNADO"
    Else
        MSFlexGrid1.TextMatrix(lin, 7) = TABLA!destino
    End If
    
    If IsNull(TABLA!cuit) Then
        MSFlexGrid1.TextMatrix(lin, 8) = "-"
    Else
        MSFlexGrid1.TextMatrix(lin, 8) = TABLA!cuit
    End If

    If TABLA!importe < "0" Then
        With MSFlexGrid1
        num = CDbl(.TextMatrix(lin, 2)) * -1
        .TextMatrix(lin, 2) = Format(num, "currency")
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 1
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 2
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 3
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 4
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 5
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 6
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 7
        .CellBackColor = &HFF00&
        '.CellFontBold = True
        .Col = 8
        .CellBackColor = &HFF00&
        End With
     End If
    If TABLA!rechazado = "-1" Then
        With MSFlexGrid1
        .Row = lin
        .Col = 0
        .CellBackColor = &HFF&
        .Col = 1
        .CellBackColor = &HFF&
        .Col = 2
        .CellBackColor = &HFF&
        .Col = 3
        .CellBackColor = &HFF&
        .Col = 4
        .CellBackColor = &HFF&
        .Col = 5
        .CellBackColor = &HFF&
        .Col = 6
        .CellBackColor = &HFF&
        .Col = 7
        .CellBackColor = &HFF&
        .Col = 8
        .CellBackColor = &HFF&
        End With
    End If
    If TABLA!rechazado = 0 Then
    
        importe = CDbl(importe) + CDbl(MSFlexGrid1.TextMatrix(lin, 2))
    End If
    TABLA.MoveNext
Loop
Label3 = importe
TABLA.Close
End If
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
L = Val(Text2) & "/" & Val(Text3) & "/" & Val(Text4)
Label9 = L
End If
End Sub

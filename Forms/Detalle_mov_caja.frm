VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Detalle_mov_caja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de los moviemientos de caja"
   ClientHeight    =   7665
   ClientLeft      =   2400
   ClientTop       =   1380
   ClientWidth     =   10575
   Icon            =   "Detalle_mov_caja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10575
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Por fecha"
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
      Left            =   6120
      Picture         =   "Detalle_mov_caja.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Listar por fecha"
      Top             =   960
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
      Left            =   7320
      Picture         =   "Detalle_mov_caja.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00 ""€"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
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
      Height          =   420
      Left            =   6840
      TabIndex        =   4
      Top             =   6960
      Width           =   3255
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
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
      Format          =   22216705
      CurrentDate     =   41058
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
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
      Format          =   22216705
      CurrentDate     =   40909
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4935
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   9615
      _ExtentX        =   16960
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "SALDO:"
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
      TabIndex        =   8
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
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
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE DE CAJA"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Detalle_mov_caja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call ALTAGRID
LISTADO = "select * from mov_caja order by r_interno desc"
TABLA.Open LISTADO, conexion_BD
saldo = 0
saldo1 = 0
resta = 0

    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!detalle
        MSFlexGrid1.TextMatrix(lin, 4) = TABLA!r_interno
        
        If IsNull(TABLA!ingreso) Then
            MSFlexGrid1.TextMatrix(lin, 2) = Format(0, "currency")
        Else
            MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!ingreso, "currency")
        End If
        
        If IsNull(TABLA!egreso) Then
            MSFlexGrid1.TextMatrix(lin, 3) = Format(0, "currency")
        Else
            MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!egreso, "currency")
        End If
        
        If TABLA!egreso <= 0 And TABLA!ingreso <= 0 Then
            MSFlexGrid1.Row = lin
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
            lin = lin - 1
        Else
            saldo = CDbl(saldo) + TABLA!ingreso
            saldo1 = CDbl(saldo1) + TABLA!egreso
        End If
        TABLA.MoveNext
    Loop
TABLA.Close
resta = saldo - saldo1

If resta < 0 Then
    Text4.BackColor = &HFF&
    Text4.ForeColor = &HFFFFFF
    
Else
    Text4.BackColor = vbGreen
    Text4.ForeColor = vbBlack
    
End If
'Text4 = resta
Text4 = Format(resta, "currency")
End Sub

Private Sub Command3_Click()
Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)


LISTADO = "select * from mov_caja where fecha >= " + CStr(inicio) + " and fecha <=  " + CStr(final) + " order by r_interno desc"
TABLA.Open LISTADO, conexion_BD

MSFlexGrid1.Clear
Call ALTAGRID
saldo = 0
saldo1 = 0
resta = 0

    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid1.TextMatrix(lin, 1) = TABLA!detalle
        MSFlexGrid1.TextMatrix(lin, 4) = TABLA!r_interno
        
        If IsNull(TABLA!ingreso) Then
            MSFlexGrid1.TextMatrix(lin, 2) = Format(0, "currency")
        Else
            MSFlexGrid1.TextMatrix(lin, 2) = Format(TABLA!ingreso, "currency")
        End If
        
        If IsNull(TABLA!egreso) Then
            MSFlexGrid1.TextMatrix(lin, 3) = Format(0, "currency")
        Else
            MSFlexGrid1.TextMatrix(lin, 3) = Format(TABLA!egreso, "currency")
        End If
        'saldo = CDbl(saldo) + TABLA!ingreso
        'saldo1 = CDbl(saldo1) + TABLA!egreso
        'resta = saldo - saldo1
        TABLA.MoveNext
    Loop
TABLA.Close
'If resta < 0 Then
'    Text4.BackColor = &HFF&
'    Text4.ForeColor = &HFFFFFF
'
'Else
'    Text4.BackColor = vbGreen
'    Text4.ForeColor = vbBlack
'
'End If
'Text4 = Format(resta, "currency")


End Sub

Private Sub Form_Load()
'DTPicker1 = Date
DTPicker2 = Date
Call ALTAGRID
LISTADO = "select * from mov_caja order by r_interno desc"
TABLA.Open LISTADO, conexion_BD
saldo = 0
saldo1 = 0
resta = 0

With MSFlexGrid1
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = MSFlexGrid1.Rows + 1
        .TextMatrix(lin, 0) = TABLA!fecha
        .TextMatrix(lin, 1) = TABLA!detalle
        If IsNull(TABLA!r_interno) Then
            .TextMatrix(lin, 4) = 0
        Else
            .TextMatrix(lin, 4) = TABLA!r_interno
        End If
        If IsNull(TABLA!ingreso) Then
            .TextMatrix(lin, 2) = Format(0, "currency")
            
        Else
            .TextMatrix(lin, 2) = Format(TABLA!ingreso, "currency")
        End If
        
        If IsNull(TABLA!egreso) Then
            .TextMatrix(lin, 3) = Format(0, "currency")
        Else
            .TextMatrix(lin, 3) = Format(TABLA!egreso, "currency")
        End If
        
        ing = .TextMatrix(lin, 2)
        egr = .TextMatrix(lin, 3)
        
        If egr <= 0 And ing <= 0 Then
            .Row = lin
            .RemoveItem (MSFlexGrid1.Row)
            lin = lin - 1
        Else
            saldo = CDbl(saldo) + ing
            saldo1 = CDbl(saldo1) + egr
        End If
        TABLA.MoveNext
    Loop
TABLA.Close
resta = saldo - saldo1
End With
If resta < 0 Then
    Text4.BackColor = &HFF&
    Text4.ForeColor = &HFFFFFF
    
Else
    Text4.BackColor = vbGreen
    Text4.ForeColor = vbBlack
    
End If
'Text4 = Format(saldo1, "currency")
Text4 = Format(resta, "currency")
End Sub
Private Sub ALTAGRID()

With MSFlexGrid1
.FixedCols = 0
.Cols = 5
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "FECHA"
.TextMatrix(0, 1) = "DETALLE"
.TextMatrix(0, 2) = "INGRESO"
.TextMatrix(0, 3) = "EGRESO"
.TextMatrix(0, 4) = "REMITO"

.ColWidth(0) = 1400
.ColWidth(1) = 4900
.ColWidth(2) = 1400
.ColWidth(3) = 1400
.ColWidth(4) = 0
End With
End Sub

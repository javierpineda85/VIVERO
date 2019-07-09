VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Listado_cheques_por_mes 
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   300
   ClientTop       =   1410
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10950
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   4095
      Left            =   5520
      TabIndex        =   7
      Top             =   2400
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7223
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
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      DisabledPicture =   "Listado_cheques_por_mes.frx":0000
      DragIcon        =   "Listado_cheques_por_mes.frx":058A
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
      Picture         =   "Listado_cheques_por_mes.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   1440
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
      Format          =   61472769
      CurrentDate     =   41179
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
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
      Format          =   61472769
      CurrentDate     =   40909
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4095
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7223
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
   Begin VB.Label Label3 
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
      Left            =   6240
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Buscar entre fechas: Desde"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Listados de Cheques por Mes"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "Listado_cheques_por_mes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Dim inicio As Long
Dim final As Long

inicio = CDate(Me.DTPicker1.value)
final = CDate(Me.DTPicker2.value)

SQL = "select * from salecheque where vto >= " + CStr(inicio) + " and vto <=" + CStr(final) + " order by vto"
TABLA.Open SQL, conexion_BD

Call CARGA_LISTA

With MSFlexGrid1
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
    mes = TABLA!vto
    mes2 = Month(mes)
    año = Year(mes)
    mes3 = MonthName(mes2) & " " & año

    .TextMatrix(lin, 0) = mes3

    If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
        If TABLA!importe < 0 Then

            cta1 = TABLA!importe * -1
            .TextMatrix(lin, 2) = cta1
            cta1 = CDbl(.TextMatrix(lin - 1, 2)) + CDbl(.TextMatrix(lin, 2))
         Else
            cta1 = 0
            .TextMatrix(lin, 2) = TABLA!importe
            cta1 = CDbl(.TextMatrix(lin - 1, 2)) + CDbl(.TextMatrix(lin, 2))
        End If

        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        .TextMatrix(lin, 1) = 0
        .TextMatrix(lin, 2) = Format(cta1, "currency")
        .TextMatrix(lin, 3) = 0

    Else
        If TABLA!importe < 0 Then

            cta2 = TABLA!importe * -1
            .TextMatrix(lin, 2) = Format(cta2, "currency")
            .TextMatrix(lin, 1) = 0
            .TextMatrix(lin, 3) = 0
         Else
            
            cta2 = TABLA!importe
            .TextMatrix(lin, 2) = Format(cta2, "currency")
            .TextMatrix(lin, 1) = 0
            .TextMatrix(lin, 3) = 0

        End If

        cta1 = 0
        cta2 = 0

                
    End If

    TABLA.MoveNext
    Loop
    TABLA.Close

End With

'For i = 0 To MSFlexGrid1.Row

SQL = "select * from entracheque where fecha_vto >= " + CStr(inicio) + " and fecha_vto <=" + CStr(final) + " order by fecha_vto"
TABLA.Open SQL, conexion_BD
Call CARGA_LISTA2

With MSFlexGrid2
'.Rows = lin + 1
lin = 0
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
    mes = TABLA!fecha_vto
    mes2 = Month(mes)
    año = Year(mes)
    mes3 = MonthName(mes2) & " " & año

    .TextMatrix(lin, 0) = mes3

    If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
        If TABLA!importe < 0 Then

            cta1 = TABLA!importe * -1
            .TextMatrix(lin, 1) = cta1
            cta1 = CDbl(.TextMatrix(lin - 1, 1)) + CDbl(.TextMatrix(lin, 1))
         Else
            cta1 = 0
            .TextMatrix(lin, 1) = TABLA!importe
            cta1 = CDbl(.TextMatrix(lin - 1, 1)) + CDbl(.TextMatrix(lin, 1))
        End If

        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
        .TextMatrix(lin, 1) = Format(cta1, "currency")
        .TextMatrix(lin, 2) = 0
        .TextMatrix(lin, 3) = 0

    Else
        If TABLA!importe < 0 Then

            cta2 = TABLA!importe * -1
            .TextMatrix(lin, 1) = Format(cta2, "currency")
            .TextMatrix(lin, 2) = 0
            .TextMatrix(lin, 3) = 0
         Else

            cta2 = TABLA!importe
            .TextMatrix(lin, 1) = Format(cta2, "currency")
            .TextMatrix(lin, 2) = 0
            .TextMatrix(lin, 3) = 0

        End If

        cta1 = 0
        cta2 = 0
                
    End If
    TABLA.MoveNext
    Loop
    TABLA.Close

End With
'Next
End Sub

Private Sub CARGA_LISTA()
With MSFlexGrid1
.FixedCols = 1
.Cols = 4
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "VENCIMIENTO"
.TextMatrix(0, 1) = "CARTERA"
.TextMatrix(0, 2) = "EMITIDOS"
.TextMatrix(0, 3) = "SUBTOTAL"


.ColWidth(0) = 2000
.ColWidth(1) = 0
.ColWidth(2) = 2000
.ColWidth(3) = 0
'.ColWidth(4) = 2000

End With
End Sub
Private Sub CARGA_LISTA2()

With MSFlexGrid2
.FixedCols = 1
.Cols = 4
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "VENCIMIENTO"
.TextMatrix(0, 1) = "CARTERA"
.TextMatrix(0, 2) = "EMITIDOS"
.TextMatrix(0, 3) = "SUBTOTAL"


.ColWidth(0) = 2000
.ColWidth(1) = 2000
.ColWidth(2) = 0
.ColWidth(3) = 0
End With
End Sub




Private Sub Form_Load()
DTPicker2 = Date
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Detalles_ctas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles de las Ctas Ctes"
   ClientHeight    =   9510
   ClientLeft      =   2460
   ClientTop       =   810
   ClientWidth     =   11115
   Icon            =   "Detalles_ctas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   11115
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   10080
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
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
      Left            =   9840
      Picture         =   "Detalles_ctas.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3600
      Width           =   975
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
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   9255
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
         Left            =   7800
         TabIndex        =   11
         Top             =   360
         Width           =   735
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
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         X1              =   3120
         X2              =   3120
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   5520
         X2              =   5520
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line4 
         X1              =   7320
         X2              =   7320
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   9000
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   9000
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   120
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line8 
         X1              =   9000
         X2              =   9000
         Y1              =   240
         Y2              =   1320
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
         Left            =   3360
         TabIndex        =   7
         Top             =   840
         Width           =   2055
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
         Left            =   5640
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
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
         Left            =   7440
         TabIndex        =   5
         Top             =   840
         Width           =   1455
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
         Left            =   5880
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Pagos realizados"
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
         TabIndex        =   9
         Top             =   360
         Width           =   2175
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
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   3015
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   480
      TabIndex        =   0
      Top             =   3600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10186
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pagos Realizados"
      TabPicture(0)   =   "Detalles_ctas.frx":0B14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Facturas"
      TabPicture(1)   =   "Detalles_ctas.frx":0B30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(4)=   "Label15"
      Tab(1).Control(5)=   "MSFlexGrid2"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Retenciones Realizadas"
      TabPicture(2)   =   "Detalles_ctas.frx":0B4C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSFlexGrid3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   3735
         Left            =   3600
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   5175
         Begin VB.CommandButton Command2 
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
            Left            =   2160
            Picture         =   "Detalles_ctas.frx":0B68
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   2760
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
            Height          =   2415
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   4260
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   4935
         Left            =   -72960
         TabIndex        =   13
         Top             =   600
         Width           =   4815
         _ExtentX        =   8493
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4575
         Left            =   -74760
         TabIndex        =   2
         Top             =   600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   8070
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4935
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
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
      Begin VB.Label Label15 
         Caption         =   "Facturas sin afectar"
         Height          =   255
         Left            =   -69600
         TabIndex        =   18
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -70080
         TabIndex        =   17
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Facturas Afectadas a un pago"
         Height          =   255
         Left            =   -73080
         TabIndex        =   16
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -73440
         TabIndex        =   15
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "REFERENCIAS:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   5400
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   5535
      Left            =   1440
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   9375
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
         Height          =   4815
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8493
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
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Detalle General de las Ctas Ctes."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Detalles_ctas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
' llamar a una rutina para armnar el flex grid sin que se vea

Call IMPRIMIR

End Sub
Private Sub IMPRIMIR()

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
Printer.Print Tab(15); "DETALLE GENERAL DE CTA. CTE. EN "; Label17
Printer.Print Tab(15); ""
Printer.Print Tab(15); ""
Printer.FontBold = True
Printer.Print Tab(10); Label2; Tab(45); Label3; Tab(67); Label4; Tab(82); Label5
Printer.Print Tab(10); Label6; Tab(45); Label7; Tab(67); Label8; Tab(82); Label9
Printer.FontBold = False
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Tab(15); ""
Printer.Print Tab(10); "Pagos Realizados"
Printer.Print Tab(15); ""
Printer.FontSize = 8
Printer.FontBold = False
Printer.Print Tab(12); "FECHA"; Tab(27); "DETALLE"; Tab(65); "MONTO"; Tab(85); "Nº DE REG."; Tab(100); "RETENCION";
Printer.Print Tab(10); "======================================================================================================"


For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    fecha = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 0)
    detalle = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 1)
    monto = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 2)
    n_reg = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3)
    
    Printer.Print Tab(10); fecha; Tab(25); detalle; Tab(65); monto; Tab(85); n_reg; Tab(100); monto_ret
Next
If sValor = "PROVEEDORES" Then
    For i = 1 To MSFlexGrid3.Rows - 1
        MSFlexGrid3.Row = i
        fecha = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 0)
        monto_ret = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 1)
        n_reg = Me.MSFlexGrid3.TextMatrix(Me.MSFlexGrid3.Row, 2)
    
        Printer.Print Tab(10); fecha; Tab(85); n_reg; Tab(100); monto_ret

    Next
End If
Printer.Print Tab(10); "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"


Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Tab(15); ""
Printer.Print Tab(10); "Facturas ingresadas"
Printer.Print Tab(15); ""
Printer.FontSize = 8
Printer.FontBold = False
Printer.Print Tab(12); "FECHA"; Tab(27); "DETALLE"; Tab(65); "MONTO"; Tab(85); "FACTURA Nº"; Tab(103); "VENCIMIENTO"; Tab(125); "ESTADO";
Printer.Print Tab(10); "======================================================================================================"


For i = 1 To MSFlexGrid2.Rows - 1
    MSFlexGrid2.Row = i
    fecha = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 0)
    factura = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 1)
    detalle = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 2)
    monto = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 3)
    vence = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 4)
    estado = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 5)
    
Printer.Print Tab(10); fecha; Tab(25); detalle; Tab(65); monto; Tab(85); factura; Tab(103); vence; Tab(125); estado;

Next
Printer.Print Tab(10); "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

Printer.EndDoc
End Sub



Private Sub Command2_Click()
Frame1.Visible = False

End Sub

Private Sub Command3_Click()
'Frame3.Visible = True
'Call ALTALISTA4

'SQL = " SELECT entracheque.n_interno,entracheque.n_cheque,entracheque.importe,entracheque.banco,entracheque.fecha_depo,cliente,entracheque.fecha_vto,mov_cheques.destino,entracheque.rechazado FROM entracheque LEFT JOIN mov_cheques ON entracheque.n_cheque=mov_cheques.n_cheque ORDER BY entracheque.fecha_vto"
'TABLA.Open SQL, conexion_BD


If sValor = "CLIENTES" Then

    SQL = "SELECT mov_clientes.pago,clientes_a_pagar.cliente,monto FROM clientes_a_pagar LEFT JOIN mov_clientes ON clientes_a_pagar.cliente=mov_clientes.cliente ORDER BY mov_clientes.cliente"
    TABLA.Open SQL, conexion_BD
    
    Call ALTALISTA4

    Do While Not TABLA.EOF
        lin = lin + 1
        With MSFlexGrid5
        .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!cliente
        
        If .TextMatrix(lin, 0) <> .TextMatrix(lin - 1, 0) Then          ''' si el clientes es <> al anterior
                                                                        ''' inserta pago y monto
            If IsNull(TABLA!pago) Then          '' validamos los valores nulos
                .TextMatrix(lin, 1) = "0"
            Else
                .TextMatrix(lin, 1) = TABLA!pago
                pag = 0

            End If
            
            If IsNull(TABLA!monto) Then         '' validamos los valores nulos
                .TextMatrix(lin, 3) = 0
            Else
                .TextMatrix(lin, 3) = TABLA!monto
                mon = 0

            End If
                
        Else
            If IsNull(TABLA!pago) Then                                  ''' el clientes es igual al anterior
                'pag = CDbl(pag) + 0                                     ''' debemos cargar en acumuladores los
                .TextMatrix(lin, 1) = "0"                              ''' valores de pago y monto
            Else
                'pag = CDbl(pag) + TABLA!pago
                'pag = CDbl(.TextMatrix(lin - 1, 1)) + TABLA!pago
                .TextMatrix(lin, 1) = TABLA!pago
                pag = CDbl(.TextMatrix(lin - 1, 1)) + CDbl(.TextMatrix(lin, 1))
                .TextMatrix(lin - 1, 1) = pag

            End If
            
            If IsNull(TABLA!monto) Then
                'mon = CDbl(mon) + 0
                .TextMatrix(lin, 3) = 0
            Else
                mon = CDbl(mon) + TABLA!monto
                .TextMatrix(lin, 3) = mon

            End If
            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            
        End If

        TABLA.MoveNext
    End With
    Loop
    
    TABLA.Close
    
Else
End If
Frame3.Visible = True
End Sub
Private Sub ALTALISTA4()
MSFlexGrid5.FixedCols = 0
MSFlexGrid5.Cols = 5
MSFlexGrid5.FixedRows = 1
MSFlexGrid5.Rows = 2
MSFlexGrid5.Clear
MSFlexGrid5.TextMatrix(0, 0) = "NOMBRE"
MSFlexGrid5.TextMatrix(0, 1) = "PAGOS"
MSFlexGrid5.TextMatrix(0, 2) = "RETENCIONES"
MSFlexGrid5.TextMatrix(0, 3) = "FACTURA"
MSFlexGrid5.TextMatrix(0, 4) = "SALDO"

MSFlexGrid5.ColWidth(0) = 3000
MSFlexGrid5.ColWidth(1) = 1500
MSFlexGrid5.ColWidth(2) = 1500
MSFlexGrid5.ColWidth(3) = 1500
MSFlexGrid5.ColWidth(4) = 1500



End Sub

Private Sub Form_Load()
Label7 = ""
Label8 = ""
Label6 = datos
suma = 0
tot2 = 0
Label14 = 0
Label17 = sValor
Select Case sValor
Case "PROVEEDORES"
    Call PAGOS_REALIZADOS
    Call A_PAGAR
    Call RETENCION
    suma = CDbl(Label7) + CDbl(Label14)
    Label7 = Format(suma, "currency")
    
Case "CLIENTES"
    Call PAGOS_REALIZADOS
    Call A_PAGAR
    MSFlexGrid3.Visible = False
    
    
Case "PERSONAL"
    MSFlexGrid3.Visible = False
End Select

If Label7 = "" Then
    Label7 = "0"
End If
If Label8 = "" Then
    Label8 = "0"
End If
resta = Label7 - Label8
Label9 = Format(resta, "currency")

End Sub
Private Sub PAGOS_REALIZADOS()

With MSFlexGrid1
Select Case sValor
Case "PROVEEDORES"
    j = "select * from mov_proveedor where proveedor= '" & datos & "' order by r_interno asc"
    TABLA.Open j, conexion_BD
    
    Call alta_prove
    lin = 0
    tot2 = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!fecha
        If IsNull(TABLA!detalle) Then
            .TextMatrix(lin, 1) = ""
        Else
            .TextMatrix(lin, 1) = TABLA!detalle
        End If
        .TextMatrix(lin, 2) = TABLA!pago
        .TextMatrix(lin, 3) = TABLA!r_interno
        
        If .TextMatrix(lin, 3) = .TextMatrix(lin - 1, 3) Then ''''' VER Q ACA ES LIN,3 Y ABAJO ES LIN,2
            
            .TextMatrix(lin, 2) = TABLA!pago
            cta = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
            .Row = lin
            .RemoveItem (MSFlexGrid1.Row)
            .TextMatrix(lin - 1, 2) = cta
            lin = lin - 1
           
        Else
        
        .BackColor = &HFFFFC0 'CAMBIA DE COLOR A CELESTE
    
        .TextMatrix(lin, 0) = TABLA!fecha
        If IsNull(TABLA!detalle) Then
            .TextMatrix(lin, 1) = ""
        Else
            .TextMatrix(lin, 1) = TABLA!detalle
        End If
        .TextMatrix(lin, 2) = TABLA!pago
        .TextMatrix(lin, 3) = TABLA!r_interno
        'tot2 = CDbl(tot2) + .TextMatrix(lin, 2)
        End If
        TABLA.MoveNext
    Loop
    TABLA.Close

    pagos = 0
    .Row = lin
    For i = 1 To .Row '- 1
    pagos = CDbl(.TextMatrix(i, 2)) + pagos
    Next
    Label7 = Format(pagos, "currency")
    
Case "CLIENTES"
    j = "select * from mov_clientes where cliente= '" & datos & "' order by r_interno asc"
    TABLA.Open j, conexion_BD

    Call alta_prove
    lin = 0
    tot2 = 0
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = MSFlexGrid1.Rows + 1
        .TextMatrix(lin, 0) = TABLA!fecha
        If IsNull(TABLA!detalle) Then
            .TextMatrix(lin, 1) = ""
        Else
            .TextMatrix(lin, 1) = TABLA!detalle
        End If
        '.TextMatrix(lin, 2) = TABLA!pago
        .TextMatrix(lin, 3) = TABLA!r_interno
        
        If .TextMatrix(lin, 3) = .TextMatrix(lin - 1, 3) Then
            .TextMatrix(lin, 2) = TABLA!pago
            cta = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
            .Row = lin
            .RemoveItem (.Row)
            .TextMatrix(lin - 1, 2) = cta
            lin = lin - 1
           
        Else
        
        .BackColor = &HFFFFC0 'CAMBIA DE COLOR A CELESTE
               
        .TextMatrix(lin, 0) = TABLA!fecha
        If IsNull(TABLA!detalle) Then
            .TextMatrix(lin, 1) = ""
        Else
            .TextMatrix(lin, 1) = TABLA!detalle
        End If
        
        .TextMatrix(lin, 2) = TABLA!pago
        .TextMatrix(lin, 3) = TABLA!r_interno
        tot2 = CDbl(tot2) + .TextMatrix(lin, 2)
        End If
        
        TABLA.MoveNext
    Loop
    TABLA.Close
    .Row = lin
    
    For i = 1 To .Row '- 1
    pagos = CDbl(.TextMatrix(i, 2)) + pagos
    Next
    
    Label7 = Format(pagos, "currency")
End Select
End With
End Sub
Private Sub A_PAGAR()

Select Case sValor

Case "PROVEEDORES"
    K = "select * from prove_a_pagar where proveedor= '" & datos & "' order by fecha asc"
    TABLA.Open K, conexion_BD
    
    Call Prove_a_pagar
    lin = 0
    tot = 0
    tot1 = 0
    tot3 = 0
    Do While Not TABLA.EOF

        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!factura
        
        If IsNull(TABLA!detalle) Then
            MSFlexGrid2.TextMatrix(lin, 2) = ""
        Else
            MSFlexGrid2.TextMatrix(lin, 2) = TABLA!detalle
        End If
        
        MSFlexGrid2.TextMatrix(lin, 3) = TABLA!monto
        MSFlexGrid2.TextMatrix(lin, 4) = TABLA!fecha_vto
        
        If TABLA!verdadero = 0 Then
            With MSFlexGrid2
            .Row = lin
            .TextMatrix(lin, 5) = "IMPAGA"
            .Col = 0
            .CellBackColor = vbRed
            .Col = 1
            .CellBackColor = vbRed
            .Col = 2
            .CellBackColor = vbRed
            .Col = 3
            .CellBackColor = vbRed
            .Col = 4
            .CellBackColor = vbRed
            
            'tot = CDbl(tot) + TABLA!monto
            
            End With

        
        Else
            With MSFlexGrid2
            
            .Row = lin
            .TextMatrix(lin, 5) = "CANCELADA"
            .Col = 0
            .CellBackColor = &HFFFFC0
            .Col = 1
            .CellBackColor = &HFFFFC0
            .Col = 2
            .CellBackColor = &HFFFFC0
            .Col = 3
            .CellBackColor = &HFFFFC0
            .Col = 4
            .CellBackColor = &HFFFFC0
            
            End With
            
    
        End If
        tot = CDbl(tot) + TABLA!monto
        TABLA.MoveNext
    Loop
    TABLA.Close
    
    Label8 = Format(tot, "currency")
    
Case "CLIENTES"
    K = "select * from clientes_a_pagar where cliente= '" & datos & "' order by fecha asc"
    TABLA.Open K, conexion_BD

    Call Prove_a_pagar
    lin = 0
    tot = 0
    tot1 = 0
    tot3 = 0
    Do While Not TABLA.EOF

        lin = lin + 1
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        MSFlexGrid2.TextMatrix(lin, 0) = TABLA!fecha
        MSFlexGrid2.TextMatrix(lin, 1) = TABLA!factura
        
        If IsNull(TABLA!detalle) Then
            MSFlexGrid2.TextMatrix(lin, 2) = ""
        Else
            MSFlexGrid2.TextMatrix(lin, 2) = TABLA!detalle
        End If
        
        MSFlexGrid2.TextMatrix(lin, 3) = TABLA!monto
        MSFlexGrid2.TextMatrix(lin, 4) = TABLA!fecha_vto
        
        If TABLA!verdadero = 0 Then
            With MSFlexGrid2
            .Row = lin
            .TextMatrix(lin, 5) = "IMPAGA"
            .Col = 0
            .CellBackColor = vbRed
            .Col = 1
            .CellBackColor = vbRed
            .Col = 2
            .CellBackColor = vbRed
            .Col = 3
            .CellBackColor = vbRed
            .Col = 4
            .CellBackColor = vbRed
            
            'tot = CDbl(tot) + TABLA!monto
            
            End With

        
        Else
            With MSFlexGrid2
            .Row = lin
            .TextMatrix(lin, 5) = "CANCELADA"
            .Col = 0
            .CellBackColor = &HFFFFC0
            .Col = 1
            .CellBackColor = &HFFFFC0
            .Col = 2
            .CellBackColor = &HFFFFC0
            .Col = 3
            .CellBackColor = &HFFFFC0
            .Col = 4
            .CellBackColor = &HFFFFC0
            
            End With
            
    
        End If
        tot = CDbl(tot) + TABLA!monto
        TABLA.MoveNext
    Loop
    TABLA.Close
    
    Label8 = Format(tot, "currency")
End Select
End Sub
Private Sub RETENCION()

a = "select * from retencion where destino='" & datos & "'order by r_interno asc"
TABLA.Open a, conexion_BD
'If TABLA!destino = "" Then
'    TABLA.Close
'    MSFlexGrid3.Visible = False
'Else
Call alta_retencion
MSFlexGrid3.Visible = True
lin = 0
tot4 = 0
With MSFlexGrid3
Do While Not TABLA.EOF
    lin = lin + 1
    .Rows = .Rows + 1
    .TextMatrix(lin, 0) = TABLA!fecha
    .TextMatrix(lin, 1) = TABLA!importe
    .TextMatrix(lin, 2) = TABLA!r_interno
    
    If TABLA!importe = 0 Then
        .Row = lin
        .RemoveItem (.Row)
        lin = lin - 1
    Else
        Label14 = CDbl(Label14) + TABLA!importe
    End If
    TABLA.MoveNext
Loop
TABLA.Close
End With
'End If
End Sub
Private Sub alta_retencion()

With MSFlexGrid3
.FixedCols = 0
.Cols = 3
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "FECHA"
.TextMatrix(0, 1) = "IMPORTE"
.TextMatrix(0, 2) = "Nº REG"

.ColWidth(0) = 2000
.ColWidth(1) = 1500
.ColWidth(2) = 1000
End With
End Sub
Private Sub alta_prove()

With MSFlexGrid1
.FixedCols = 0
.Cols = 4
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "FECHA"
.TextMatrix(0, 1) = "DETALLE"
.TextMatrix(0, 2) = "MONTO"
.TextMatrix(0, 3) = "Nº REG"
'.TextMatrix(0, 4) = "PAGO EN"

.ColWidth(0) = 1500
.ColWidth(1) = 3800
.ColWidth(2) = 1500
.ColWidth(3) = 1500
'.ColWidth(4) = 1500
End With
End Sub
Private Sub Prove_a_pagar()

'proveedores a pagar

With MSFlexGrid2
.FixedCols = 0
.Cols = 6
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "FECHA"
.TextMatrix(0, 1) = "FACTURA"
.TextMatrix(0, 2) = "DETALLE"
.TextMatrix(0, 3) = "MONTO"
.TextMatrix(0, 4) = "VENCE"
.TextMatrix(0, 5) = "ESTADO"

.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 2500
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 0

End With
End Sub



Private Sub MSFlexGrid1_Click()
Label16 = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Row, 3) ' Trae el numero del registro

If sValor = "CLIENTES" Then

    bus = "select * from mov_clientes where r_interno='" & Label16 & "'"
    TABLA.Open bus, conexion_BD

    Frame1.Visible = True
    Call FLEX4
    
    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
        MSFlexGrid4.TextMatrix(lin, 0) = TABLA!tipo
        MSFlexGrid4.TextMatrix(lin, 1) = TABLA!n_cheque
        MSFlexGrid4.TextMatrix(lin, 2) = TABLA!pago
        
        TABLA.MoveNext
    Loop
    
    TABLA.Close
    
    'For i = 2 To MSFlexGrid4.Rows - 1
    '    SQL = "select * from entracheque where n_cheque= " & Val(MSFlexGrid4.TextMatrix(i, 1)) & ""
    '    TABLA.Open SQL, conexion_BD
    '    Do While Not TABLA.EOF
     '       'i = i + 1
    '        MSFlexGrid4.TextMatrix(i, 3) = TABLA!fecha_vto
    '    Loop
    '    TABLA.Close
    'Next
Else

    bus = " select * from mov_proveedor where r_interno='" & Label16 & "'"
    TABLA.Open bus, conexion_BD

    Frame1.Visible = True
    Call FLEX4
    Do While Not TABLA.EOF
        lin = lin + 1
        MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
        MSFlexGrid4.TextMatrix(lin, 0) = TABLA!tipo
        MSFlexGrid4.TextMatrix(lin, 1) = TABLA!n_cheque
        MSFlexGrid4.TextMatrix(lin, 2) = TABLA!pago
        
        TABLA.MoveNext
    Loop
    
    TABLA.Close
End If
End Sub
Private Sub FLEX4()
With MSFlexGrid4
.FixedCols = 0
.Cols = 3
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "FORMA DE PAGO"
.TextMatrix(0, 1) = "Nº CHEQUE"
.TextMatrix(0, 2) = "MONTO"
'.TextMatrix(0, 3) = "VENCIMIENTO"


.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
'.ColWidth(3) = 1500
End With
End Sub

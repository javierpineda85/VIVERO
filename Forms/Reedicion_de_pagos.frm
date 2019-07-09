VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Reedicion_de_pagos 
   Caption         =   "Gestión de Modificación de Pagos"
   ClientHeight    =   8805
   ClientLeft      =   1950
   ClientTop       =   555
   ClientWidth     =   9270
   Icon            =   "Reedicion_de_pagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   9270
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar"
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
      Left            =   7080
      Picture         =   "Reedicion_de_pagos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   3000
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   4575
      Begin VB.ComboBox Combo2 
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
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label21 
         Caption         =   "Seleccione el Nombre/Razón Social que desea buscar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por:"
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
      Left            =   480
      TabIndex        =   16
      Top             =   1680
      Width           =   2535
      Begin VB.OptionButton OpPro 
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
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton OpPer 
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
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton OpCli 
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
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   2175
      End
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
      Left            =   7080
      Picture         =   "Reedicion_de_pagos.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   6975
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
         Left            =   5040
         TabIndex        =   28
         Top             =   2160
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1575
         Left            =   1080
         TabIndex        =   27
         Top             =   3240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2778
         _Version        =   393216
         BackColor       =   16777152
         BackColorBkg    =   -2147483633
         Appearance      =   0
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
         Left            =   1800
         TabIndex        =   7
         Top             =   960
         Width           =   1815
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
         TabIndex        =   6
         Top             =   960
         Width           =   1815
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
         Left            =   1200
         TabIndex        =   5
         Top             =   1560
         Width           =   5415
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4800
         Picture         =   "Reedicion_de_pagos.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guardar Cambios"
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
         Left            =   4800
         Picture         =   "Reedicion_de_pagos.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
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
         Left            =   5880
         Picture         =   "Reedicion_de_pagos.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4080
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   2160
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
         Format          =   22282241
         CurrentDate     =   41177
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   6000
         TabIndex        =   32
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   22282241
         CurrentDate     =   41212
      End
      Begin VB.Label Label9 
         Caption         =   "Movimiento eliminado el dia"
         Height          =   375
         Left            =   3960
         TabIndex        =   31
         Top             =   2760
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Detalle del pago:"
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
         TabIndex        =   30
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Efectivo:"
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
         TabIndex        =   29
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
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
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
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
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Remito:"
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
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre/Razón social:"
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
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
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
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5055
      Left            =   480
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   8295
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
         Left            =   3600
         Picture         =   "Reedicion_de_pagos.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4080
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3615
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6376
         _Version        =   393216
         BackColor       =   16777152
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Modificación de Pagos"
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
      Left            =   2040
      TabIndex        =   26
      Top             =   480
      Width           =   5295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Reedicion_de_pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As String
Private Sub FACTURAS_GRID()
With MSFlexGrid2

.FixedCols = 0
.Cols = 4
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "Nº REMITO"
.TextMatrix(0, 1) = "FECHA"
.TextMatrix(0, 2) = "MONTO"
.TextMatrix(0, 3) = "DETALLE"

.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 4000
End With
End Sub
Private Sub FACTURAS_GRID1()
With MSFlexGrid2

.FixedCols = 0
.Cols = 4
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "Nº REMITO"
.TextMatrix(0, 1) = "FECHA"
.TextMatrix(0, 2) = "ADELANTO"
.TextMatrix(0, 3) = "DETALLE"

.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 4000
End With
End Sub
Private Sub Alta_cheque()
With MSFlexGrid1

.FixedCols = 0
.Cols = 3
.FixedRows = 1
.Rows = 2
.Clear
.TextMatrix(0, 0) = "Nº CHEQUE"
.TextMatrix(0, 1) = "MONTO"
.TextMatrix(0, 2) = "TIPO"
'.TextMatrix(0, 3) = "DETALLE"

.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 0
'.ColWidth(3) = 4000
End With
End Sub

Private Sub Combo2_Click()
Select Case op

Case "PROVEEDORES"
    SQL = "select * from mov_proveedor where proveedor='" & Combo2 & "'"
    TABLA.Open SQL, conexion_BD
    Call FACTURAS_GRID
    With MSFlexGrid2
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!r_interno
        .TextMatrix(lin, 1) = TABLA!fecha
        
        If IsNull(TABLA!pago) Then
            .TextMatrix(lin, 2) = "SIN DATOS"
        Else
            .TextMatrix(lin, 2) = TABLA!pago
        End If
        
        If IsNull(TABLA!detalle) Then
            .TextMatrix(lin, 3) = "SIN DATOS"
        Else
            .TextMatrix(lin, 3) = TABLA!detalle
        End If
        
        If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
        
            suma = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 2) = suma
        
        End If
        
        TABLA.MoveNext
    Loop
    TABLA.Close
    End With
    
Case "CLIENTES"
    
    SQL = "select * from mov_clientes where cliente='" & Combo2 & "'"
    TABLA.Open SQL, conexion_BD
    Call FACTURAS_GRID
    With MSFlexGrid2
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        .TextMatrix(lin, 0) = TABLA!r_interno
        .TextMatrix(lin, 1) = TABLA!fecha
        
        If IsNull(TABLA!pago) Then
            .TextMatrix(lin, 2) = "SIN DATOS"
        Else
            .TextMatrix(lin, 2) = TABLA!pago
        End If
        
        If IsNull(TABLA!detalle) Then
            .TextMatrix(lin, 3) = "SIN DATOS"
        Else
            .TextMatrix(lin, 3) = TABLA!detalle
        End If
        
        If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
        
            suma = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 2) = suma
        
        End If
        
        TABLA.MoveNext
    Loop
    TABLA.Close
    End With
    
Case "PERSONAL"
    SQL = "select * from mov_rrhh where nombre_rrhh='" & Combo2 & "'"
    TABLA.Open SQL, conexion_BD
    
    Call FACTURAS_GRID1
    With MSFlexGrid2
    
    Do While Not TABLA.EOF
        lin = lin + 1
        .Rows = .Rows + 1
        '.TextMatrix(lin, 0) = TABLA!r_interno
        .TextMatrix(lin, 1) = TABLA!fecha_mod
        
        If IsNull(TABLA!adelanto) Then
            .TextMatrix(lin, 2) = "SIN DATOS"
        Else
            .TextMatrix(lin, 2) = TABLA!adelanto
        End If
        
        If IsNull(TABLA!detalle) Then
            .TextMatrix(lin, 3) = "SIN DATOS"
        Else
            .TextMatrix(lin, 3) = TABLA!detalle
        End If
        
        If .TextMatrix(lin, 0) = .TextMatrix(lin - 1, 0) Then
        
            suma = CDbl(.TextMatrix(lin, 2)) + CDbl(.TextMatrix(lin - 1, 2))
            .Row = lin
            .RemoveItem (.Row)
            lin = lin - 1
            .TextMatrix(lin, 2) = suma
        
        End If
        
        TABLA.MoveNext
    Loop
    TABLA.Close
    End With

End Select
Frame4.Visible = True
Command2.Visible = True
End Sub

Private Sub Command2_Click()
Label2 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Frame3.Visible = False
Frame4.Visible = False
Command2.Visible = False

End Sub

Private Sub Command3_Click()
Frame3.Visible = False
Frame4.Visible = True
End Sub

Private Sub Command4_Click()
If op = "CLIENTES" Then
    With MSFlexGrid1
    
   'MODIFICA REMITO
    mod_rem = " update remito_interno set total='" & Text3 & "', detalle='" & Text4 & "',fecha='" & DTPicker1 & "' where n_remito=" & Val(Text2) & ""
    conexion_BD.Execute mod_rem
    
   'MODIFICA CAJA
    mod_caja = " update mov_caja set ingreso='" & Text1 & "' where r_interno='" & Text2 & "'"
    conexion_BD.Execute mod_caja
    
   'MODIFICA ENTRACHEQUE '''' solo se puede modificar un cheque por la modificacion de cheques y puede hacerlo solo USUA nivel 1 o 2
   'For i = 1 To .Rows - 1
   '     mod_che = " update entracheque set importe='" & .TextMatrix(i, 1) & "' where n_cheque= " & Val(.TextMatrix(i, 0)) & ""
   '     conexion_BD.Execute mod_che
   ' Next
    
   'MODIFICA EL MOVIMIENTO DEL CLIENTE

    For i = 1 To .Rows - 1
        If .TextMatrix(i, 2) = "EFECTIVO" Then
        
            mod_mov = " update mov_clientes set detalle='" & Text4 & "', fecha='" & DTPicker1 & "', pago= '" & Text1 & "'  where r_interno='" & Text2 & "' and n_cheque= " & Val(.TextMatrix(1, 0)) & ""
            conexion_BD.Execute mod_mov
        Else
       
            mod_mov = " update mov_clientes set detalle='" & Text4 & "', fecha='" & DTPicker1 & "' where r_interno='" & Text2 & "'"
            conexion_BD.Execute mod_mov
        
        End If
            
    Next
    
   End With
   
Else
    With MSFlexGrid1
    
   'MODIFICA REMITO
    mod_rem = " update remito_interno set total='" & Text3 & "', detalle='" & Text4 & "',fecha='" & DTPicker1 & "' where n_remito=" & Val(Text2) & ""
    conexion_BD.Execute mod_rem
    
   'MODIFICA CAJA
    mod_caja = " update mov_caja set egreso='" & .TextMatrix(1, 1) & "' where r_interno='" & Text2 & "'"
    conexion_BD.Execute mod_caja
    
   'MODIFICA ENTRACHEQUE SOLO SE PUEDE HACER SI ES NIVEL USUA 1 o 2 y por la edicion de cheques
   'For i = 1 To .Rows - 1
   '     mod_che = " update salecheque set importe='" & .TextMatrix(i, 1) & "' where n_cheque= " & Val(.TextMatrix(i, 0)) & ""
   '     conexion_BD.Execute mod_che
   ' Next
    
   'MODIFICA EL MOVIMIENTO DEL PROVEEDOR
        'modifica el detalle
    mod_mov = " update mov_proveedor set detalle='" & Text4 & "', fecha='" & DTPicker1 & "' where r_interno='" & Text2 & "'"
    conexion_BD.Execute mod_mov
    
        'modifica monto en efectivo
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 2) = "EFECTIVO" Then
        
            mod_mov = " update mov_proveedor set detalle='" & Text4 & "', fecha='" & DTPicker1 & "', pago= '" & Text1 & "'  where r_interno='" & Text2 & "' and n_cheque= " & Val(.TextMatrix(1, 0)) & ""
            conexion_BD.Execute mod_mov
        Else
    
            mod_mov = " update mov_proveedor set detalle='" & Text4 & "', fecha='" & DTPicker1 & "' where r_interno='" & Text2 & "'"
            conexion_BD.Execute mod_mov
        
        End If
            
    Next
    
   End With
End If
Frame3.Visible = False
Frame4.Visible = False
Label2 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command5_Click()
If op = "CLIENTES" Then
    ques = MsgBox("Esta por eliminar de manera permanente una factura, desea continuar?", vbYesNo)
    If ques = vbYes Then
        SQL = " delete * from mov_clientes where r_interno='" & Text2 & "' and cliente='" & Label2 & "'"
        conexion_BD.Execute SQL
        
        mod_rem = " update remito_interno set detalle='" & Label9 & " " & DTPicker2 & " " & usua & "' where n_remito=" & Val(Text2) & ""
        conexion_BD.Execute mod_rem
        
    End If
Else
    ques = MsgBox("Esta por eliminar de manera permanente una factura, desea continuar?", vbYesNo)
    If ques = vbYes Then
        SQL = " delete * from mov_proveedor where r_interno='" & Text2 & "' and proveedor='" & Label2 & "'"
        conexion_BD.Execute SQL
        
        mod_rem = " update remito_interno set detalle='" & Label9 & " " & DTPicker2 & " " & usua & "' where n_remito=" & Val(Text2) & ""
        conexion_BD.Execute mod_rem
    End If
End If
Frame3.Visible = False
Frame4.Visible = False
Label2 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command7_Click()
Frame4.Visible = False
End Sub



Private Sub Form_Load()
DTPicker2 = Date
End Sub

Private Sub MSFlexGrid2_Click()
Label2 = Combo2.Text
Text2 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 0)        ' nº de remito
DTPicker1 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 1)    ' fecha
Text3 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 2)        ' monto
Text4 = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.Row, 3)        ' detalle

Select Case op
    Case "CLIENTES"
    
        SQL = "select * from mov_clientes where r_interno='" & Text2 & "'"
        TABLA.Open SQL, conexion_BD
        
        Call Alta_cheque
        With MSFlexGrid1
        
        Do While Not TABLA.EOF
            lin = lin + 1
            .Rows = .Rows + 1
            .TextMatrix(lin, 0) = TABLA!n_cheque
            .TextMatrix(lin, 1) = TABLA!pago
            .TextMatrix(lin, 2) = TABLA!tipo
            If TABLA!tipo = "EFECTIVO" Then
                Text1 = TABLA!pago
            End If
            
            TABLA.MoveNext
        Loop
        TABLA.Close
        
        End With
        
    Case "PROVEEDORES"
            
        SQL = "select * from mov_proveedor where r_interno='" & Text2 & "'"
        TABLA.Open SQL, conexion_BD
        
        Call Alta_cheque
        With MSFlexGrid1
        
        Do While Not TABLA.EOF
            lin = lin + 1
            .Rows = .Rows + 1
            .TextMatrix(lin, 0) = TABLA!n_cheque
            .TextMatrix(lin, 1) = TABLA!pago
            .TextMatrix(lin, 2) = TABLA!tipo
            If TABLA!tipo = "EFECTIVO" Then
                Text1 = TABLA!pago
            End If
            TABLA.MoveNext
        Loop
        TABLA.Close
        
        End With
        
    End Select
Frame3.Visible = True
Frame4.Visible = False
End Sub



Private Sub OpCli_Click()
If OpCli = True Then
    Combo2.Clear
    op = "CLIENTES"
    SQL = "select * from clientes order by nombre_cliente"
    TABLA.Open SQL, conexion_BD
    
    Do While Not TABLA.EOF
        Combo2.AddItem TABLA!nombre_cliente
        
        TABLA.MoveNext
    Loop
    TABLA.Close
End If
Frame2.Visible = True
Frame2.BorderStyle = 0
End Sub


Private Sub OpPer_Click()
If OpPer = True Then
    op = "PERSONAL"
    Combo2.Clear
    SQL = " select * from alta_rrhh order by nombre_rrhh"
    TABLA.Open SQL, conexion_BD
    
    Do While Not TABLA.EOF
        Combo2.AddItem TABLA!nombre_rrhh
        TABLA.MoveNext
    Loop
    TABLA.Close
End If
Frame2.Visible = True
Frame2.BorderStyle = 0
End Sub

Private Sub OpPro_Click()
If OpPro = True Then
    op = "PROVEEDORES"
    Combo2.Clear
    SQL = " select * from proveedores order by nombre_prove"
    TABLA.Open SQL, conexion_BD
    
    Do While Not TABLA.EOF
        Combo2.AddItem TABLA!nombre_prove
        TABLA.MoveNext
    Loop
    TABLA.Close
End If
Frame2.Visible = True
Frame2.BorderStyle = 0

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub


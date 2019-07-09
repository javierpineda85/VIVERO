VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mov_banco_eliminarcheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar o Rechazar un Cheque"
   ClientHeight    =   8040
   ClientLeft      =   2175
   ClientTop       =   1065
   ClientWidth     =   9975
   Icon            =   "Mov_banco_eliminarcheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   9975
   Visible         =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   4575
      Left            =   6000
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         DisabledPicture =   "Mov_banco_eliminarcheque.frx":058A
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
         Left            =   120
         Picture         =   "Mov_banco_eliminarcheque.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rechazar"
         DisabledPicture =   "Mov_banco_eliminarcheque.frx":109E
         DragIcon        =   "Mov_banco_eliminarcheque.frx":1628
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
         Left            =   120
         Picture         =   "Mov_banco_eliminarcheque.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         DisabledPicture =   "Mov_banco_eliminarcheque.frx":213C
         DragIcon        =   "Mov_banco_eliminarcheque.frx":26C6
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
         Left            =   120
         Picture         =   "Mov_banco_eliminarcheque.frx":2C50
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   5055
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   4695
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
         Left            =   1800
         TabIndex        =   35
         Top             =   2160
         Width           =   2535
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
         Left            =   1800
         TabIndex        =   33
         Top             =   3960
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   4560
         Width           =   2535
         _ExtentX        =   4471
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
         CurrentDate     =   41214
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   1560
         Width           =   2535
         _ExtentX        =   4471
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
         CurrentDate     =   41214
      End
      Begin VB.TextBox Text7 
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
         Left            =   1800
         TabIndex        =   16
         ToolTipText     =   "IMPORTANTE: si se cambia el destino del cheque es importante poner el nombre tal cual esta en la cuenta corriente"
         Top             =   3960
         Width           =   2535
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
         Left            =   1800
         TabIndex        =   14
         Top             =   3360
         Width           =   2535
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
         Left            =   1800
         TabIndex        =   13
         ToolTipText     =   $"Mov_banco_eliminarcheque.frx":31DA
         Top             =   2760
         Width           =   2535
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
         Left            =   1800
         TabIndex        =   12
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Ingreso:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Destino:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Banco:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Importe:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Detalle/cliente:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Vencimiento:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Nº interno:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Nº cheque:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   9135
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "Mov_banco_eliminarcheque.frx":326B
         DragIcon        =   "Mov_banco_eliminarcheque.frx":37F5
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
         Picture         =   "Mov_banco_eliminarcheque.frx":3D7F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Recibidos"
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
         Left            =   6360
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Emitidos"
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
         Left            =   5040
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Height          =   390
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Ingrese el nº de cheque:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   7320
      TabIndex        =   36
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label17 
      Caption         =   "Label17 para buscar los r_interno del mov_prove"
      Height          =   495
      Left            =   480
      TabIndex        =   34
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "remito interno"
      Height          =   375
      Left            =   2640
      TabIndex        =   31
      Top             =   7320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "FECHA"
      Height          =   255
      Left            =   8400
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   495
      Left            =   4320
      TabIndex        =   29
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "respaldo de n_cheque"
      Height          =   495
      Left            =   6000
      TabIndex        =   28
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   255
      Left            =   7680
      TabIndex        =   27
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modificar o Rechazar un Cheque"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   7335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Mov_banco_eliminarcheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Option1 = True Then
'salecheque
    'modifica datos en salecheque
    modif = " update salecheque set n_cheque=" & Val(Text2) & ", vto='" _
    & DTPicker1 & "',detalle= '" & Text4 & "', importe='" & Text5 & "', banco='" _
    & Text6 & "',destino='" & Combo1 & "',fecha='" & DTPicker2 & "' where n_interno = " & Val(Label12) & ""
    conexion_BD.Execute modif
    
    ' modifica datos en mov_cheque
    SQL = " update mov_cheques set n_cheque=" & Val(Text2) & ",fecha_vto='" _
    & DTPicker1 & "',detalle= '" & Text4 & "', importe='" & Text5 & "',destino='" _
    & Combo1 & "',fecha_mod='" & DTPicker2 & "' where n_interno = " & Val(Label12) & ""
    conexion_BD.Execute SQL
    
    'modifica datos en remito interno
    sql2 = "select * from mov_proveedor where n_cheque=" & Val(Text2) & ""
    TABLA.Open sql2, conexion_BD
    
    Do While Not TABLA.EOF
        remito_int = TABLA!r_interno
        TABLA.MoveNext
    Loop
    TABLA.Close
    Label17 = remito_int
    
    SQL3 = " update remito_interno set detalle= '" & Text4 & ". Se modificaron datos de cheque/s',destino='" _
    & Combo1 & "',fecha='" & DTPicker2 & "' where n_remito=" & Val(Label17) & ""
    conexion_BD.Execute SQL3
    
    remito = "insert into remito_interno values (" & Val(Label16) & ",'" & Label15 & "','VIVERO SAN NICOLAS S.A.','" & Text4 & "','MODIFICACION EN CHEQUE','" & Text5 & "','" & usua & "')"
    conexion_BD.Execute remito
    
    'modifica datos en mov_prove
    SQL1 = " update mov_proveedor set detalle= '" & Text4 & "', pago='" & Text5 & "',proveedor='" _
    & Combo1 & "',fecha='" & DTPicker2 & "',r_interno= " & Val(Label16) & " where n_cheque=" & Val(Text2) & ""
    conexion_BD.Execute SQL1
    
Else
'entracheque
    'modifica los datos en entracheque
    modif = " update entracheque set n_cheque=" & Val(Text2) & ", fecha_vto='" _
    & DTPicker1 & "',cliente= '" & Combo2 & "', importe='" & Text5 & "', banco='" _
    & Text6 & "',fecha_depo='" & DTPicker2 & "' where n_interno = " & Val(Label12) & ""
    conexion_BD.Execute modif
  
    ' modifica datos en mov_cheque
    SQL = " update mov_cheques set n_cheque=" & Val(Text2) & ",fecha_vto='" _
    & DTPicker1 & "',detalle= '" & Combo2 & "', importe='" & Text5 & "',destino='" _
    & Combo1 & "',fecha_mod='" & DTPicker2 & "' where n_interno = " & Val(Label12) & ""
    conexion_BD.Execute SQL
   
    'modifica datos en remito interno
    sql2 = "select * from mov_clientes where n_cheque=" & Val(Text2) & ""
    TABLA.Open sql2, conexion_BD
    
    Do While Not TABLA.EOF
        remito_int = TABLA!r_interno
        TABLA.MoveNext
    Loop
    TABLA.Close
    Label17 = remito_int
    
    SQL3 = " update remito_interno set detalle= '" & Combo2 & ". Se modificaron datos de cheque/s',destino='" _
    & Combo1 & "',fecha='" & DTPicker2 & "' where n_remito=" & Val(Label17) & ""
    conexion_BD.Execute SQL3
    
    remito = "insert into remito_interno values (" & Val(Label16) & ",'" & Label15 & "','VIVERO SAN NICOLAS S.A.','" & Combo2 & "','MODIFICACION EN CHEQUE','" & Text5 & "','" & usua & "')"
    conexion_BD.Execute remito
    
    'modifica datos en mov_clientes
    SQL1 = " update mov_clientes set detalle= '" & Combo2 & "', pago='" & Text5 & "',cliente='" _
    & Combo2 & "',fecha='" & DTPicker2 & "',r_interno= " & Val(Label16) & " where n_cheque=" & Val(Text2) & ""
    conexion_BD.Execute SQL1
    
End If
MsgBox "Los datos han sido modificados correctamente!"
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Label12 = ""
End Sub

Private Sub Command2_Click()
If Option1 = True Then
    Label6 = "Detalle"
    
    SQL = "select * from proveedores order by nombre_prove"
    TABLA.Open SQL, conexion_BD
    Do While Not TABLA.EOF
        Combo1.AddItem TABLA!nombre_prove
        TABLA.MoveNext
    Loop
    TABLA.Close
    
    buscar = "select * from salecheque where n_cheque= " & Val(Text1) & ""
    TABLA.Open buscar, conexion_BD
    
    Do While Not TABLA.EOF
        ninterno = TABLA!n_interno
        ncheque = TABLA!n_cheque
        vto = TABLA!vto
        detalle = TABLA!detalle
        importe = TABLA!importe
        banco = TABLA!banco
        destino = TABLA!destino
        fecha = TABLA!fecha
        recha = TABLA!rechazado

        TABLA.MoveNext
    
    Loop
    
    TABLA.Close
    
    Text2 = ncheque
    Label13 = ncheque 'copia de seguridad, por si se modifica el nº antes de rechazarlo
    Label12 = ninterno
    DTPicker1 = vto
    Text4 = detalle
    Text5 = importe
    Text6 = banco
    Combo1 = destino
    DTPicker2 = fecha
    Label18 = recha
    
    If Label18 = -1 Then
        Label18 = "EL CHEQUE YA HA SIDO RECHAZADO"
    Else
        Label18 = ""
    End If
        
    Frame2.Visible = True
    Frame3.Visible = True
    
    Combo1.Visible = True
    Combo2.Visible = False
    Text7.Visible = False
    
Else
    Label6 = "Cliente"
    
    SQL = "select * from clientes order by nombre_cliente"
    TABLA.Open SQL, conexion_BD
    Do While Not TABLA.EOF
        Combo2.AddItem TABLA!nombre_cliente
        TABLA.MoveNext
    Loop
    TABLA.Close
    
    buscar = "select * from entracheque where n_cheque= " & Val(Text1) & ""
    TABLA.Open buscar, conexion_BD
    
    Do While Not TABLA.EOF
        ninterno = TABLA!n_interno
        ncheque = TABLA!n_cheque
        vto = TABLA!fecha_vto
        detalle = TABLA!cliente
        importe = TABLA!importe
        banco = TABLA!banco
        fecha = TABLA!fecha_depo
        Label13 = "Clientes"
        recha = TABLA!rechazado
        
        TABLA.MoveNext
    
    Loop
    TABLA.Close
    
    Text2 = ncheque
    Label13 = ncheque 'copia de seguridad, por si se modifica el nº antes de rechazarlo
    Label12 = ninterno
    DTPicker1 = vto
    Combo2 = detalle
    Text5 = importe
    Text6 = banco
    Text7 = destino
    DTPicker2 = fecha
    
    Label18 = recha
    
    If Label18 = -1 Then
        Label18 = "EL CHEQUE YA HA SIDO RECHAZADO"
    Else
        Label18 = ""
    End If
    
    Frame2.Visible = True
    Frame3.Visible = True

    Combo1.Visible = False
    Combo2.Visible = True
    Text7.Visible = True

End If


End Sub

Private Sub Command3_Click()
' al rechazar el cheque debe generar un saldo por el monto
' del cheque en la cuenta del proveedor o cliente.
' pero no debe figurar ni ingresado ni por ingresar al banco

buscar = " select * from mov_proveedor where n_cheque=" & Val(Text2) & ""
TABLA.Open buscar, conexion_BD

If TABLA.EOF Then
    TABLA.Close
    
    buscar = " select * from mov_clientes where n_cheque=" & Val(Text2) & ""
    TABLA.Open buscar, conexion_BD
    
    If TABLA!n_cheque = Val(Text2) Then
        TABLA.Close
        Label14 = CDbl(Text5) * -1
        ' insertar pago en negativo en los clientes
        recha = " insert into mov_clientes values ('" & Combo2 & "','CHEQUE RECHAZADO','" & Label14 & "'," & Val(Text2) & ",'" & Label15 & "','" & Label16 & "','CHEQUE')"
        conexion_BD.Execute recha
        
        Label11 = "-1"
        modif = "update entracheque set rechazado= " & Val(Label11) & " where n_interno = " & Val(Label12) & ""
        conexion_BD.Execute modif
        
        ' Esta 2 veces el combo2 por que tanto el origen como el destino es lo mismo
        remito = "insert into remito_interno values (" & Val(Label16) & ",'" & Label15 & "','" & Combo2 & "','" & Combo2 & "','CHEQUE RECHAZADO','" & Label14 & "','" & usua & "')"
        conexion_BD.Execute remito
        
        Borrar = "delete * from enbanco where n_cheque= " & Val(Text2) & ""
        conexion_BD.Execute Borrar
        
        'inserta moviemiento de cheque para dar destino al cheque en el listado de historicos
        mov_che = "insert into mov_cheques values (" & Val(Label12) & "," & Val(Text2) & ",'" & Label15 & "','" & DTPicker1 & "','" & Combo2 & "','" & Combo2 & "','" & Text5 & "','CHEQUE RECHAZADO')"
        conexion_BD.Execute mov_che
        
    End If
    

Else
    TABLA.Close
    
    Label14 = CDbl(Text5) * -1
    
    recha = "insert into mov_proveedor values ('" & Text4 & "', 'CHEQUE RECHAZADO','" & Label14 & "'," & Val(Text2) & ",'" & Label15 & "','" & Label16 & "','CHEQUE CH')"
    conexion_BD.Execute recha
    
    Label11 = "-1"
    modif_rech = "update salecheque set rechazado= " & Val(Label11) & " where n_interno = " & Val(Label12) & ""
    conexion_BD.Execute modif_rech
    
    remito = "insert into remito_interno values (" & Val(Label16) & ",'" & Label15 & "','" & Text7 & "','" & Text7 & "','CHEQUE RECHAZADO','" & Label14 & "','" & usua & "')"
    conexion_BD.Execute remito
    
    Borrar = "delete * from enbanco where n_cheque= " & Val(Text2) & ""
    conexion_BD.Execute Borrar
    
    'inserta moviemiento de cheque para dar destino al cheque en el listado de historicos
    mov_che = "insert into mov_cheques values (" & Val(Label12) & "," & Val(Text2) & ",'" & Label15 & "','" & DTPicker1 & "','" & Text4 & "','" & Text4 & "','" & Text5 & "','CHEQUE RECHAZADO')"
    conexion_BD.Execute mov_che
    
End If


MsgBox "El cheque ha sido rechazado!"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Label12 = ""

End Sub

Private Sub Command4_Click()


res = MsgBox(" Esta seguro que desea borrar de manera permanente los datos del cheque?", vbYesNo)
If res = vbYes Then

    If Option1 = True Then
        'salecheque
        'elimina datos en salecheque
        Borrar = " delete * from salecheque where n_interno = " & Val(Label12) & ""
        conexion_BD.Execute Borrar
    
        ' elimina datos en mov_cheque
        SQL = " delete * from mov_cheques where n_interno = " & Val(Label12) & ""
        conexion_BD.Execute SQL
      
        'elimina datos en remito interno
        sql2 = "select * from mov_proveedor where n_cheque=" & Val(Text2) & ""
        TABLA.Open sql2, conexion_BD
   
        Do While Not TABLA.EOF
            remito_int = TABLA!r_interno
            TABLA.MoveNext
        Loop
        TABLA.Close
        Label17 = remito_int
    
        remito = " update remito_interno set detalle= '" & Text4 & ". Se elimino un/os cheque/s &',destino='" _
        & Combo1 & "',fecha='" & DTPicker2 & "' where n_remito=" & Val(Label17) & ""
        conexion_BD.Execute remito
   
        'elimina datos en mov_prove
        SQL1 = " delete * from mov_proveedor where n_cheque=" & Val(Text2) & ""
        conexion_BD.Execute SQL1
    
        'elimina del banco
        Borrar = " delete * from enbanco where n_interno = " & Val(Label12) & ""
        conexion_BD.Execute Borrar
    
        MsgBox "Los datos han sido eliminados correctamente!"
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Text6 = ""
        Text7 = ""
        Text8 = ""
        Label12 = ""
    
    Else
        'entracheque
        'elimina los datos en entracheque
        modif = " delete * from entracheque where n_interno = " & Val(Label12) & ""
        conexion_BD.Execute modif

        ' modifica datos en mov_cheque
        SQL = " delete * from mov_cheques where n_interno = " & Val(Label12) & ""
        conexion_BD.Execute SQL
    
        'elimina datos en remito interno
        sql2 = "select * from mov_clientes where n_cheque=" & Val(Text2) & ""
        TABLA.Open sql2, conexion_BD
   
        Do While Not TABLA.EOF
            remito_int = TABLA!r_interno
            TABLA.MoveNext
        Loop
        TABLA.Close
        Label17 = remito_int
   
        remito = " update remito_interno set detalle= '" & Text4 & ". Se elimino un/os cheque/s &',destino='" _
        & Combo1 & "',fecha='" & DTPicker2 & "' where n_remito=" & Val(Label17) & ""
        conexion_BD.Execute remito
    
        'elimina datos en mov_clientes
        SQL1 = " delete * from mov_clientes where n_cheque=" & Val(Text2) & ""
        conexion_BD.Execute SQL1
    
        'elimina del banco
        Borrar = " delete * from enbanco where n_interno = " & Val(Label12) & ""
        conexion_BD.Execute Borrar
    
        MsgBox "Los datos han sido eliminados correctamente!"
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Text6 = ""
        Text7 = ""
        Text8 = ""
        Label12 = ""

    End If
End If
End Sub

Private Sub Form_Load()
Frame1.BorderStyle = 0
Frame2.BorderStyle = 0
Frame3.BorderStyle = 0
Label15 = Date

remito = "select max(n_remito) from remito_interno"
TABLA.Open remito, conexion_BD
Label16 = TABLA.Fields(0) + 1
TABLA.Close

End Sub

Private Sub Option1_Click()
If Option1 = True Then
    tablacheque = "salecheque"
    Text7.Enabled = True

Else
    tablacheque = "entracheque"
    Text7.Enabled = False

End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
    tablacheque = "entracheque"
    Text7.Enabled = False

Else
    tablacheque = "salecheque"
    Text7.Enabled = True

End If
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mov_cheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de cheques"
   ClientHeight    =   7830
   ClientLeft      =   3030
   ClientTop       =   915
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mov_cheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   8055
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar"
      Height          =   735
      Left            =   5520
      Picture         =   "Mov_cheques.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "IMPORTANTE! Borra todo lo que contienen las listas!"
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Guardar"
      Height          =   735
      Left            =   6480
      Picture         =   "Mov_cheques.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List7 
      Height          =   1425
      Left            =   6960
      TabIndex        =   40
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agregar"
      Height          =   735
      Left            =   4560
      Picture         =   "Mov_cheques.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
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
      Height          =   2295
      Left            =   240
      TabIndex        =   20
      Top             =   5280
      Width           =   7215
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
         Height          =   1260
         Left            =   0
         TabIndex        =   26
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
         Height          =   1260
         Left            =   1080
         TabIndex        =   25
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
         Height          =   1260
         Left            =   2280
         TabIndex        =   24
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
         Height          =   1260
         Left            =   3360
         TabIndex        =   23
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
         Height          =   1260
         Left            =   4440
         TabIndex        =   22
         Top             =   840
         Width           =   1215
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
         Height          =   1260
         Left            =   5640
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "CUIT"
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
         Left            =   6000
         TabIndex        =   32
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Vencimiento"
         Height          =   255
         Left            =   4560
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
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
         Left            =   3600
         TabIndex        =   30
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
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
         Left            =   2400
         TabIndex        =   29
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label11 
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
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label10 
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
         TabIndex        =   27
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar Cambios"
      Height          =   735
      Left            =   6480
      Picture         =   "Mov_cheques.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Mov_cheques.frx":1BB2
      Left            =   1320
      List            =   "Mov_cheques.frx":1BB4
      TabIndex        =   4
      Top             =   3360
      Width           =   4815
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
      Left            =   1080
      MaxLength       =   11
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
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
      CurrentDate     =   41018
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   4080
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
      CurrentDate     =   41018
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
      Left            =   4440
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
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
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
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
      Height          =   405
      Left            =   4920
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label22 
      Caption         =   "L22=convertit"
      Height          =   735
      Left            =   240
      TabIndex        =   39
      Top             =   4560
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label21 
      Caption         =   "Pago a cuenta corriente"
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
      Left            =   6720
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label20 
      Caption         =   "SISTEMA=VIVERO O STELLA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "L19=total"
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
      Left            =   3240
      TabIndex        =   36
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label18 
      Caption         =   "L18=fecha"
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
      Left            =   6840
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "L17=rem inter"
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
      TabIndex        =   34
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label16 
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
      Left            =   1800
      TabIndex        =   33
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "CUIT:"
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
      TabIndex        =   14
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha vencimiento:"
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
      Left            =   3000
      TabIndex        =   13
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha deposito:"
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
      Left            =   2280
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      TabIndex        =   16
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Left            =   3480
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Importe:  $"
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
      TabIndex        =   17
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Nº:"
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
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Interno:"
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
      TabIndex        =   19
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHEQUES RECIBIDOS"
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
      Width           =   4935
   End
End
Attribute VB_Name = "Mov_cheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opt As String

Private Sub Command1_Click()

If Text2.Text = Val(List2.List(i)) Then
    MsgBox "Este cheque ya ha sido ingresado"
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text2.SetFocus

Else

    List1.AddItem (Label16)
    List2.AddItem (Text2)
    List3.AddItem (Text3)
    List4.AddItem (Text4)
    List5.AddItem (DTPicker1)
    List6.AddItem (Text5)

    m = MsgBox("Desea cargar otro cheque?", vbYesNo, "VIVERO SAN NICOLAS")

    If m = vbYes Then
        Label16 = Label16 + 1
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Text2.SetFocus
        Command2.Visible = True
    Else
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Command2.Visible = True
    End If

End If

End Sub

Private Sub Command2_Click()
T = 0
For j = 0 To List1.ListCount - 1
    T = CDbl(T) + List3.List(j)
    
B = "select * from entracheque where n_cheque= " & Val(List2.List(j)) & "    "
TABLA.Open B, conexion_BD
Do While Not TABLA.EOF
    List7.AddItem TABLA!n_cheque
    TABLA.MoveNext
Loop
TABLA.Close

If Val(List2.List(j)) = Val(List7.List(j)) Then
    MsgBox "Algunos de los cheques ya han sido ingresados"
End If
Next


Label19 = T
txtmon = CDbl(Label19)
Call CONVERTIR
Label22 = txtmonl
Call IMPRIMIR
''' EN STELLA CAMBIA LABEL 20 X LABEL 23
rint = "insert into remito_interno values (" & Val(Label17) & ",'" & Label18 & "','" & Combo1 & "','" & Label20 & "','" & Label21 & "','" & Label19 & "','" & usua & "')"
conexion_BD.Execute rint

For i = 0 To List1.ListCount - 1

    che = "insert into entracheque values (" & Val(List1.List(i)) & "," & Val(List2.List(i)) & ",'" & List3.List(i) & "','" & Combo1 & "','" & List4.List(i) & "','" & List5.List(i) & "','" & Label18 & "'," & 0 & ")"
    conexion_BD.Execute che
    
    produ = "insert into mov_clientes values ('" & Combo1 & "','" & Label21 & "','" & List3.List(i) & "'," & Val(List2.List(i)) & ",'" & Label18 & "'," & Val(Label17) & ",'CHEQUE')"
    conexion_BD.Execute produ
Next

Mov_cheques.Refresh
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Label16 = Val(Label16) + 1

List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
Combo1.Clear


End Sub


Private Sub Command4_Click()

Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
Label7.Visible = False
'Label9.Visible = False
'Label10.Visible = False
'Text5.Visible = False
'Text6.Visible = False
DTPicker2.Visible = False

interno = "select max(n_interno) from entracheque"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Label16 = 10000
Else
    Label16 = TABLA.Fields(0) + 1
End If
TABLA.Close
End Sub



Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
Label18 = Date
Label20 = SISTEMA
interno = "select max(n_interno) from entracheque"
TABLA.Open interno, conexion_BD
If TABLA.EOF = True Then
    Label16 = 10000
Else
    Label16 = TABLA.Fields(0) + 1
End If
TABLA.Close

    d = "select * from clientes order by nombre_cliente"
    TABLA.Open d, conexion_BD
    Combo1.Clear
    Do While Not TABLA.EOF
        
        Combo1.AddItem TABLA!nombre_cliente
        TABLA.MoveNext
    Loop
    TABLA.Close

End Sub

Private Sub IMPRIMIR()
remito = "select max(n_remito) from remito_interno"
TABLA.Open remito, conexion_BD
Label17 = TABLA.Fields(0) + 1
TABLA.Close
Printer.CurrentX = 30
Printer.CurrentY = 80
Printer.FontSize = 10
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); "Nº de recibo: "; Label17; 'label 21= n_interno
Printer.Print Tab(110); Label18
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí/mos de VIVERO SAN NICOLAS SA."
'Printer.Print Tab(15); " Recibí/mos de STELLA DAVIRE."
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos $ "; Label19;
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " ( "; Label22; " )."
Printer.Print Tab(15); ""
Printer.Print Tab(15); " En concepto de: "; Text1;  '" por la factura Nº: "; Text2
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Según el siguiente detalle:"
Printer.Print Tab(15); ""
Printer.Print Tab(15); '"Efectivo en $ "; Text4
Printer.Print Tab(15); '""
Printer.Print Tab(15); "INTERNO Nº "; Tab(35); "CHEQUE Nº"; Tab(55); "IMPORTE"; Tab(75); "BANCO"; Tab(98); "VENCIMIENTO"; 'Tab(95); "CUIT:"; Tab(110)
Printer.Print Tab(15); "=============================================================================="

For a = 0 To List1.ListCount - 1
    List1.ListIndex = a: n_interno = List1.List(a)
    List2.ListIndex = a: n_cheque = List2.List(a)
    List3.ListIndex = a: importe = List3.List(a)
    List4.ListIndex = a: banco = List4.List(a)
    List5.ListIndex = a: fecha = List5.List(a)
    List6.ListIndex = a: CUIT = List6.List(a)

Printer.Print Tab(15); " VT "; n_interno; Tab(35); n_cheque; Tab(55); importe; Tab(75); banco; Tab(98); fecha; 'Tab(95); CUIT; Tab(110)
Next

Printer.Print Tab(15); "-------------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Monto total en $ "; Label19
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(15); ""
Printer.Print Tab(70); Combo1.Text
Printer.Print Tab(15); ""
Printer.Print Tab(15); "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.Print Tab(15); "Nº de recibo: "; Label17; 'label 21= n_interno
Printer.Print Tab(110); Label18
Printer.Print Tab(10); ""
Printer.Print Tab(10); ""
Printer.CurrentX = 100
Printer.Print Tab(15); " Recibí/mos de VIVERO SAN NICOLAS SA."
'Printer.Print Tab(15); " Recibí/mos de STELLA DAVIRE."
Printer.Print Tab(15); ""
Printer.Print Tab(15); " la cantidad de pesos $ "; Label19;
'Printer.Print Tab(15); ""
'Printer.Print Tab(15); " ( "; Label22; " )."
Printer.Print Tab(15); ""
Printer.Print Tab(15); " En concepto de: "; Text1;  '" por la factura Nº: "; Text2
Printer.Print Tab(15); ""
Printer.Print Tab(15); "Según el siguiente detalle:"
Printer.Print Tab(15); ""
Printer.Print Tab(15); '"Efectivo en $ "; Text4
Printer.Print Tab(15); '""
Printer.Print Tab(15); "INTERNO Nº "; Tab(35); "CHEQUE Nº"; Tab(55); "IMPORTE"; Tab(75); "BANCO"; Tab(98); "VENCIMIENTO"; 'Tab(95); "CUIT:"; Tab(110)
Printer.Print Tab(15); "=============================================================================="

For a = 0 To List1.ListCount - 1
    List1.ListIndex = a: n_interno = List1.List(a)
    List2.ListIndex = a: n_cheque = List2.List(a)
    List3.ListIndex = a: importe = List3.List(a)
    List4.ListIndex = a: banco = List4.List(a)
    List5.ListIndex = a: fecha = List5.List(a)
    List6.ListIndex = a: CUIT = List6.List(a)

Printer.Print Tab(15); " VT "; n_interno; Tab(35); n_cheque; Tab(55); importe; Tab(75); banco; Tab(98); fecha; 'Tab(95); CUIT; Tab(110)
Next

Printer.Print Tab(15); "-------------------------------------------------------------------------------------------------------------------------------------"
Printer.Print Tab(15); ""
Printer.Print Tab(15); " Monto total en $ "; Label19
Printer.Print Tab(60); "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
Printer.Print Tab(15); ""
Printer.Print Tab(70); Combo1.Text
Printer.Print Tab(15); ""
Printer.Print Tab(15); "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

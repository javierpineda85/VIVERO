Attribute VB_Name = "Module1"
Public conexion_BD As New ADODB.Connection
Public TABLA As New ADODB.Recordset
Global txtmon As String
Global txtmonl As String
Global sValor As String
Global usua As String
Global NIVEL_U As String
Global SISTEMA As String
Global SISTEMA_DIR As String
Global datos As String
Global datos1 As String
Global datos2 As String
Global datos3 As String
Global datos4 As String
Global datos5 As String
Global datos6 As String
Global datos7 As String
Global datos8 As String
Global datos9 As String
Global datos10 As String
Global datos11 As String
Global datos12 As String
Global datos13 As String



Sub abrir()
    conexion_BD.ConnectionString = App.Path + "\viverodb.mdb"
    SISTEMA = "VIVERO SAN NICOLAS S.A."
    SISTEMA_DIR = "RUTA PROVINCIAL 60 S/N. JUNIN, MENDOZA."
    
    'conexion_BD.ConnectionString = App.Path + "\stelladb.mdb"
    'SISTEMA = "STELLA DAVIRE"
    'SISTEMA_DIR = ""
    
    'conexion_BD.ConnectionString = App.Path + "\viverodb.mdb"
    'SISTEMA = "WINE CONCEPT"
    'SISTEMA_DIR = "RUTA PROVINCIAL 60 S/N. JUNIN, MENDOZA"
    
    'conexion_BD.ConnectionString = App.Path + "\\Datos\Public\Paquete\Support\viverodb.mdb"
    
    conexion_BD.Provider = "Microsoft.Jet.OLEDB.4.0;Data Source = " & conexion_BD.ConnectionString & ";" & " Jet OLEDB:Database Password=" & 2586

    'TABLA.CursorType = adOpenKeyset
    'TABLA.LockType = adLockOptimistic
    conexion_BD.Open
End Sub

Sub cerrar()
    conexion_BD.Close
End Sub
Public Sub CONVERTIR()
Dim sValor As String, siValor As Single

Dim i, r As Single

sValor = txtmon

txtmon = Val(Replace(sValor, ",", "."))

i = Int(Val(txtmon))

r = CInt((txtmon - i) * 100)

Num2Text (Val(txtmon))

If r > 0 Then

txtmonl = txtmonl & " CON " + Num2Text(r) + " CENTAVOS"

End If

'txtmon = Format(txtmon, "currency")

End Sub

Public Function Num2Text(ByVal value As Double) As String

value = Int(value)

   Select Case value

       Case 0: Num2Text = "CERO"

       Case 1: Num2Text = "UN"

       Case 2: Num2Text = "DOS"

       Case 3: Num2Text = "TRES"

       Case 4: Num2Text = "CUATRO"

       Case 5: Num2Text = "CINCO"

       Case 6: Num2Text = "SEIS"

       Case 7: Num2Text = "SIETE"

       Case 8: Num2Text = "OCHO"

       Case 9: Num2Text = "NUEVE"

       Case 10: Num2Text = "DIEZ"

       Case 11: Num2Text = "ONCE"

       Case 12: Num2Text = "DOCE"

       Case 13: Num2Text = "TRECE"

       Case 14: Num2Text = "CATORCE"

       Case 15: Num2Text = "QUINCE"

       Case Is < 20: Num2Text = "DIECI" & Num2Text(value - 10)

       Case 20: Num2Text = "VEINTE"

       Case Is < 30: Num2Text = "VEINTI" & Num2Text(value - 20)

       Case 30: Num2Text = "TREINTA"

       Case 40: Num2Text = "CUARENTA"

       Case 50: Num2Text = "CINCUENTA"

       Case 60: Num2Text = "SESENTA"

       Case 70: Num2Text = "SETENTA"

       Case 80: Num2Text = "OCHENTA"

       Case 90: Num2Text = "NOVENTA"

       Case Is < 100: Num2Text = Num2Text(Int(value \ 10) * 10) & " Y " & Num2Text(value Mod 10)

       Case 100: Num2Text = "CIEN"

       Case Is < 200: Num2Text = "CIENTO " & Num2Text(value - 100)

       Case 200, 300, 400, 600, 800: Num2Text = Num2Text(Int(value \ 100)) & "CIENTOS"

       Case 500: Num2Text = "QUINIENTOS"

       Case 700: Num2Text = "SETECIENTOS"

       Case 900: Num2Text = "NOVECIENTOS"

       Case Is < 1000: Num2Text = Num2Text(Int(value \ 100) * 100) & " " & Num2Text(value Mod 100)

       Case 1000: Num2Text = "MIL"

       Case Is < 2000: Num2Text = "MIL " & Num2Text(value Mod 1000)

       Case Is < 1000000: Num2Text = Num2Text(Int(value \ 1000)) & " MIL"

           If value Mod 1000 Then Num2Text = Num2Text & " " & Num2Text(value Mod 1000)

       Case 1000000: Num2Text = "UN MILLON"

       Case Is < 2000000: Num2Text = "UN MILLON " & Num2Text(value Mod 1000000)

       Case Is < 1000000000000#: Num2Text = Num2Text(Int(value / 1000000)) & " MILLONES "

           If (value - Int(value / 1000000) * 1000000) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000) * 1000000)

       Case 1000000000000#: Num2Text = "UN BILLON"

       Case Is < 2000000000000#: Num2Text = "UN BILLON " & Num2Text(value - Int(value / 1000000000000#) * 1000000000000#)

       Case Else: Num2Text = Num2Text(Int(value / 1000000000000#)) & " BILLONES"

           If (value - Int(value / 1000000000000#) * 1000000000000#) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000000000#) * 1000000000000#)

   End Select

If value = 1 Then

 txtmonl = Num2Text + " PESO"

Else

 txtmonl = Num2Text + " PESOS"

End If

End Function

Public Function IMPRIMIR()

End Function

Public Function PERMISOS()
'If usua = "Javier Pineda" Then

End Function

VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Menú Principal"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15255
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":10B67
   Begin VB.Menu altas 
      Caption         =   "Altas"
      Begin VB.Menu alta_produ 
         Caption         =   "De Clientes"
      End
      Begin VB.Menu alta_empresas 
         Caption         =   "De Empresas"
         Visible         =   0   'False
      End
      Begin VB.Menu alta_rrhh 
         Caption         =   "De Personal"
      End
      Begin VB.Menu alta_provee 
         Caption         =   "De Proveedores"
      End
      Begin VB.Menu alta_usua 
         Caption         =   "De Usuarios"
      End
   End
   Begin VB.Menu movimientos 
      Caption         =   "Movimientos"
      Begin VB.Menu mov_clien 
         Caption         =   "De Clientes"
      End
      Begin VB.Menu mov_personal 
         Caption         =   "De Personal"
      End
      Begin VB.Menu mov_prove 
         Caption         =   "De Proveedores"
      End
      Begin VB.Menu mov_tar 
         Caption         =   "De Tarjeta"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu lista_clientes 
         Caption         =   "De Clientes"
      End
      Begin VB.Menu listado_facturasxvencer 
         Caption         =   "De Facturas a Vencer"
         Visible         =   0   'False
      End
      Begin VB.Menu lista_ctacte 
         Caption         =   "De Saldos de Ctas Ctes"
      End
      Begin VB.Menu lista_perso 
         Caption         =   "De Personal"
      End
      Begin VB.Menu lista_prove 
         Caption         =   "De Proveedores"
      End
      Begin VB.Menu lista_retencion 
         Caption         =   "De Retenciones"
      End
      Begin VB.Menu lista_tarjeta 
         Caption         =   "De Tarjeta"
      End
   End
   Begin VB.Menu banco 
      Caption         =   "Banco"
      Begin VB.Menu ingresocheque 
         Caption         =   "Ingreso de Cheques"
         Visible         =   0   'False
      End
      Begin VB.Menu listado_2 
         Caption         =   "Listado"
         Begin VB.Menu listaxingresar 
            Caption         =   "Listado de Cheques por Ingresar"
         End
         Begin VB.Menu Listaxingresados 
            Caption         =   "Listado de Cheques Ingresados"
         End
         Begin VB.Menu lista_caja 
            Caption         =   "Listado de Caja"
         End
      End
      Begin VB.Menu mov_1 
         Caption         =   "Movimientos"
         Begin VB.Menu mov_caj 
            Caption         =   "Movimiento de  Caja"
         End
         Begin VB.Menu mov_ban 
            Caption         =   "Movimiento de Banco"
         End
      End
   End
   Begin VB.Menu cheques 
      Caption         =   "Cheques"
      Begin VB.Menu mov_cheque 
         Caption         =   "Movimiento de  Cheques"
         Visible         =   0   'False
      End
      Begin VB.Menu list_cheques 
         Caption         =   "Listado de Cheques"
         Begin VB.Menu cartera 
            Caption         =   "Cartera"
         End
         Begin VB.Menu list_ch_emit 
            Caption         =   "Emitidos"
         End
         Begin VB.Menu lista_ingreso_egreso 
            Caption         =   "Ingresos/Egresos"
         End
         Begin VB.Menu hist 
            Caption         =   "Historicos"
            Begin VB.Menu list_ch_rec 
               Caption         =   "Recibidos"
            End
         End
      End
      Begin VB.Menu modifcheque 
         Caption         =   "Modificar un cheque"
      End
      Begin VB.Menu resumen_x_mes 
         Caption         =   "Resumen por Mes"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu gest_especial 
      Caption         =   "Gestiones Especiales"
      Begin VB.Menu editar_factura 
         Caption         =   "Modificación de Facturas"
      End
      Begin VB.Menu editar_pagos 
         Caption         =   "Modificación de Pagos"
      End
      Begin VB.Menu remitos 
         Caption         =   "Reimpresión de Remitos"
      End
   End
   Begin VB.Menu datos_programador 
      Caption         =   "Datos al programador"
   End
   Begin VB.Menu salir 
      Caption         =   "Salir"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_cliente_Click()
Alta_clientes.Show
End Sub

Private Sub alta_empresas_Click()
Alta_empresa.Show
End Sub

Private Sub alta_produ_Click()
Alta_clientes.Show
End Sub

Private Sub alta_provee_Click()
alta_proveedor.Show
End Sub

Private Sub alta_rrhh_Click()
Altas_personal.Show

End Sub



Private Sub alta_usua_Click()
Alta_usuarios.Show
End Sub

Private Sub cartera_Click()
Mov_cheques_cartera.Show
End Sub


Private Sub list_cheques_cartera_Click()
Listado_cheques.Show
End Sub

Private Sub editar_factura_Click()
Reedicion_facturas.Show
End Sub

Private Sub editar_pagos_Click()
Reedicion_de_pagos.Show
End Sub

Private Sub list_ch_emit_Click()
Listado_cheques_emitidos.Show
End Sub

Private Sub list_ch_rec_Click()
Listado_cheques.Show
End Sub

Private Sub lista_caja_Click()
Detalle_mov_caja.Show
End Sub

Private Sub lista_clientes_Click()
Listado_clientes.Show
End Sub

Private Sub lista_ctacte_Click()
Listado_x_ctacte.Show
End Sub

Private Sub lista_ingreso_egreso_Click()
Listado_ingreso_egreso.Show
End Sub

Private Sub lista_perso_Click()
Listado_rrhh.Show
End Sub

Private Sub lista_prove_Click()
Listado_proveedores.Show
End Sub

Private Sub lista_retencion_Click()
Listado_retenciones.Show
End Sub

Private Sub lista_tarjeta_Click()
Listado_tarjeta.Show
End Sub

Private Sub listado_facturasxvencer_Click()
Listado_fact_vencer.Show
End Sub

Private Sub Listaxingresados_Click()
Listados_cheques_en_banco.Show
End Sub

Private Sub listaxingresar_Click()
Listado_cheques_a_ingresar.Show
End Sub

Private Sub MDIForm_Load()
Call abrir
Select Case NIVEL_U

Case "1"
    'NIVEL DE PROGRAMACION. Privilegio: accede a todo el sistema y a los archivos *.vbp

Case "2"
    'NIVEL ADMINISTRADOR. Privilegio:accede a todo el sistema pero no a los archivos *.vbp
    datos_programador.Visible = False
    lista_ingreso_egreso.Visible = False
    mov_caj.Visible = True 'los movimientos los realizan desde las ctas ctes
    
Case "3"
    'NIVEL USUARIO AVANZADO. Privilegio: Accede a todo el sistema menos al modulo de cheques
    cheques.Enabled = False
    alta_usua.Enabled = False
    datos_programador.Visible = False
    editar_factura.Enabled = False
    editar_pagos.Enabled = False
    Listado_ingreso_egreso.Visible = False
    mov_caj.Visible = True 'los movimientos los realizan desde las ctas ctes
    
Case "4"
    'NIVEL USUARIO OPERARIO. Privilegio: Puede ver Listados pero no el modulo de cheques
    cheques.Enabled = False
    'banco.Enabled = False
    alta_usua.Enabled = False
    datos_programador.Visible = False
    editar_factura.Enabled = False
    editar_pagos.Enabled = False
    Listado_ingreso_egreso.Visible = False
    listado_2.Enabled = False
    'listados.Enabled = False
    mov_caj.Visible = True 'los movimientos los realizan desde las ctas ctes

Case "5"
    'NIVEL USUARIO OPERARIO 2. Privilegio: Puede ver banco, listados y cheques
    datos_programador.Visible = False
    altas.Enabled = False
    movimientos.Enabled = False
    gest_especial.Visible = False
    Listado_ingreso_egreso.Visible = False
    mov_caj.Visible = False 'los movimientos los realizan desde las ctas ctes

Case "6"
    'NIVEL USUARIO INVITADO. Privilegio: Solo accede a los listados de ctas ctes
    datos_programador.Visible = False
    altas.Enabled = False
    movimientos.Enabled = False
    banco.Enabled = False
    cheques.Enabled = False
    gest_especial.Enabled = False
    Listado_ingreso_egreso.Visible = False
    mov_caj.Visible = False 'los movimientos los realizan desde las ctas ctes
    
End Select
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Call cerrar
End Sub

Private Sub mov_banco_Click()

End Sub

Private Sub modifcheque_Click()
Mov_banco_eliminarcheque.Show
End Sub

Private Sub mov_ban_Click()
Mov_Banco.Show
End Sub

Private Sub mov_caj_Click()

Mov_caja.Show

End Sub


Private Sub mov_cheque_Click()
Mov_cheques.Show

End Sub

Private Sub mov_clien_Click()
Mov_cliente.Show
End Sub

Private Sub mov_personal_Click()
Mov_rrhh.Show
End Sub

Private Sub mov_prove_Click()
Mov_proveedores.Show
'MsgBox " EN REPARACION "
End Sub

Private Sub mov_tar_Click()
Mov_tarjeta.Show
End Sub

Private Sub remitos_Click()
Reimpresiones.Show
End Sub

Private Sub resumen_x_mes_Click()
Listado_cheques_por_mes.Show
End Sub

Private Sub salir_Click()

'Login.Show
'Unload Me

End Sub

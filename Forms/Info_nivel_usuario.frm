VERSION 5.00
Object = "Word.Document.8"; "WINWORD.EXE"
Begin VB.Form Info_nivel_usuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información de Niveles de Usuarios"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   Icon            =   "Info_nivel_usuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin WordCtl.Document Document1 
      Height          =   6810
      Left            =   120
      OleObjectBlob   =   "Info_nivel_usuario.frx":058A
      TabIndex        =   0
      Top             =   0
      Width           =   7830
   End
End
Attribute VB_Name = "Info_nivel_usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

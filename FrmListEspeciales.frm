VERSION 5.00
Begin VB.Form FrmListEspeciales 
   Caption         =   "Listado de animales con indicaciones especiales"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form5"
   ScaleHeight     =   1560
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Listado de animales con indicaciones especiales"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "FrmListEspeciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RptListado_de_Animales_con_Indicaciones_especiales.Show
End Sub


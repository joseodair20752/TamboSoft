VERSION 5.00
Begin VB.Form FrmListado_Celos 
   Caption         =   "Listado de Celos"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   1380
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Listado de Vacas en celo"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "FrmListado_Celos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RptListado_General_de_Celos.Show
End Sub

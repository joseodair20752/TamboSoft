VERSION 5.00
Begin VB.Form FRMLISTADO_ENFERMOS 
   Caption         =   "LISTADO DE ANIMALES ENFERMOS "
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Listado de animales enfermos "
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "FRMLISTADO_ENFERMOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RptListado_General_de_Enfermedades.Show
End Sub

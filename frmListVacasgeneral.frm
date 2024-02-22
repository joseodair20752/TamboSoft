VERSION 5.00
Begin VB.Form frmListVacasgeneral 
   Caption         =   "Listado de Vacas"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form5"
   ScaleHeight     =   1230
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Listado de vacas"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmListVacasgeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
RptListado_General_de_Vacas_en_Produccion.Show
End Sub

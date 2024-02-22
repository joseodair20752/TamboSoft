VERSION 5.00
Begin VB.Form frmListVaquillonas 
   Caption         =   "Listado de Vaquillonas"
   ClientHeight    =   1485
   ClientLeft      =   3570
   ClientTop       =   3495
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   1485
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000004&
      Caption         =   "Listado de Vaquillonas"
      Height          =   615
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmListVaquillonas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmListVaquillonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()


End Sub

Private Sub Command2_Click()
RptListadoVaquillonas.Show
End Sub

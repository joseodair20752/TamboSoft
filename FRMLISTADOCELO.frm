VERSION 5.00
Begin VB.Form FRMLISTADOCELO 
   Caption         =   "LISTADO DE CELOS "
   ClientHeight    =   1530
   ClientLeft      =   3390
   ClientTop       =   3975
   ClientWidth     =   5670
   LinkTopic       =   "Form5"
   ScaleHeight     =   1530
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "LISTADO DE CELOS "
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "FRMLISTADOCELO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
RptListado_General_de_Celos.Show
End Sub

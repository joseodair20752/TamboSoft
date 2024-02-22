VERSION 5.00
Begin VB.Form frmlistado_en_taque_de_Leche 
   Caption         =   "Listado del Cantidad en Tanque Primer ordeñe"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Listado de Cantidad en Tanque Segundo ordeñe"
      Height          =   855
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listado del Cantidad en Tanque Primer ordeñe"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmlistado_en_taque_de_Leche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RptListado_General_del_Primer_Ordeñe.Show
End Sub

Private Sub Command2_Click()
RptListado_General_del_Segundo_Ordeñe.Show
'frmDataEnv.Hide
End Sub

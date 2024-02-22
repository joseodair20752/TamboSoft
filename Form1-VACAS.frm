VERSION 5.00
Begin VB.Form FRMPAR1 
   Caption         =   "BUSQUEDA DE VACAS "
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "BUSQUEDA DE VACAS"
      Height          =   615
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1-VACAS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "FRMPAR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim RP As String
RP = InputBox(" INGRESE EL RP DEL ANIMAL QUE DESEE BUSCAR ", "MENU DE BUSQUEDA")
If RP = "" Then Exit Sub
With DataEnvironment1
With .rsPARAMETRORP
If .State = adStateOpen Then .Close
End With
Call .PARAMETRORP(RP)
If .rsPARAMETRORP.EOF Then
MsgBox "NO HAY NINGUN REGISTRO, POR FAVOR DIGITE DE NUEVO"
Else
RPTCARGA_DE_VACAS.Show
End If
End With

FRMPAR1.Hide

End Sub


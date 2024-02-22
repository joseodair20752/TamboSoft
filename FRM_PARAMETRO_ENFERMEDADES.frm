VERSION 5.00
Begin VB.Form FRM_PARAMETRO_ENFERMEDADES 
   Caption         =   "Busqueda de animales enfermos"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   1245
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "Busqueda de Animales Enfermos"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "FRM_PARAMETRO_ENFERMEDADES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim RP As String
RP = InputBox(" INGRESE EL RP DEL ANIMAL QUE DESEE BUSCAR ", "MENU DE BUSQUEDA")
If RP = "" Then Exit Sub
With DataEnvironment1
With .rsPARAMETRO_ENFERMEDADES
If .State = adStateOpen Then .Close
End With
Call .PARAMETRO_ENFERMEDADES(RP)
If .rsPARAMETRO_ENFERMEDADES.EOF Then
MsgBox "NO HAY NINGUN REGISTRO, POR FAVOR DIGITE DE NUEVO"
Else
RPT_CARGA_DE_ENFERMEDADES.Show
End If
End With
FRM_PARAMETRO_ENFERMEDADES.Hide
End Sub

VERSION 5.00
Begin VB.Form FRM_CELOS_PARAMETRO 
   Caption         =   "BUSQUEDA DE ANIMALES EN ESTADO DE CELO"
   ClientHeight    =   1560
   ClientLeft      =   4290
   ClientTop       =   3840
   ClientWidth     =   5580
   Icon            =   "FRM_CELOS_PARAMETRO.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   1560
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "BUSQUEDA DE ANIMALES CON CELOS"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "FRM_CELOS_PARAMETRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim RP As String
RP = InputBox(" INGRESE EL RP DEL ANIMAL QUE DESEE BUSCAR ", "MENU DE BUSQUEDA")
If RP = "" Then Exit Sub
With DataEnvironment1
With .rsPARAMETROS_CARGA_CELOS
If .State = adStateOpen Then .Close
End With
Call .PARAMETROS_CARGA_CELOS(RP)
If .rsPARAMETROS_CARGA_CELOS.EOF Then
MsgBox "NO HAY NINGUN REGISTRO, POR FAVOR DIGITE DE NUEVO"
Else
RPT_PARAMETRO_DE_CELOS.Show
End If
End With
FRM_CELOS_PARAMETRO.Hide
End Sub

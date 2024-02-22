VERSION 5.00
Begin VB.Form FRM_CONTROL_ZOMATICO 
   Caption         =   "PARAMETROS DE CONTROLES ZOMÁTICOS"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   2070
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "BUSQUEDA DE CONTROLES ZOMATICOS"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "FRM_CONTROL_ZOMATICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim RP As String
'RP = InputBox(" INGRESE EL RP DEL ANIMAL QUE DESEE BUSCAR ", "MENU DE BUSQUEDA")
'If RP = "" Then Exit Sub
'With DataEnvironment1
'With .CONTROL_SOMATICO
'If .State = adStateOpen Then .Close
'End With
'Call .PARAMETRO_ZOMATICO(RP)
'If .PARAMETRO_ZOMATICO.EOF Then
'MsgBox "NO HAY NINGUN REGISTOS, POR FAVOR DIGITE DE NUEVO"
'Else
'rpt_Zomatico.Show
'End If
'End With
'FRM_CONTROL_ZOMATICO.Hide
End Sub

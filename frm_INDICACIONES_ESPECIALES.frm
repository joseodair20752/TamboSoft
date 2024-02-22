VERSION 5.00
Begin VB.Form frm_INDICACIONES_ESPECIALES 
   Caption         =   "ANIMALES CON INDICACIONES ESPECIALES"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form3"
   ScaleHeight     =   1620
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "BUSQUEDA DE ANIMALES CON INDICACIONES ESPECIALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frm_INDICACIONES_ESPECIALES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim RP As String
RP = InputBox(" INGRESE EL RP DEL ANIMAL QUE DESEE BUSCAR ", "MENU DE BUSQUEDA")
If RP = "" Then Exit Sub
With DataEnvironment1
With .rsPARAMETRO_INDICACIONES_ESPECIALES
If .State = adStateOpen Then .Close
End With
Call .PARAMETRO_INDICACIONES_ESPECIALES(RP)
If .rsPARAMETRO_INDICACIONES_ESPECIALES.EOF Then
MsgBox "NO HAY NINGUN REGISTRO, POR FAVOR DIGITE DE NUEVO"
Else
RPT_ANIMALES_CON_INDICACIONES.Show
End If
End With
frm_INDICACIONES_ESPECIALES.Hide

End Sub

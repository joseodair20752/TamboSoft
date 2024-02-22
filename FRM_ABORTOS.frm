VERSION 5.00
Begin VB.Form FRM_ABORTOS 
   Caption         =   "BUSQUEDA DE ANIMALES CON ABORTOS"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FRM_ABORTOS.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "BUSQUEDA DE ABORTOS"
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
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "FRM_ABORTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim RP As String
RP = InputBox(" INGRESE EL RP DEL ANIMAL QUE DESEE BUSCAR ", "MENU DE BUSQUEDA")
If RP = "" Then Exit Sub
With DataEnvironment1
With .rsCARGA_DE_ABORTOS
If .State = adStateOpen Then .Close
End With
Call .PARAMETROS_ABORTOS(RP)
If .rsPARAMETROS_ABORTOS.EOF Then
MsgBox "NO HAY NINGUN REGISTRO, POR FAVOR DIGITE DE NUEVO"
Else
RPT_ABORTOS.Show
End If
End With
FRM_ABORTOS.Hide
End Sub

VERSION 5.00
Begin VB.Form FRM_SEGUNDO_ORDEÑE 
   Caption         =   "DATOS DEL SEGUNDO ORDEÑE"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   1470
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "DATOS DEL SEGUNDO ORDEÑE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "FRM_SEGUNDO_ORDEÑE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim FECHA As Date
FECHA = InputBox(" INGRESE LA FECHA DEL ORDEÑE QUE DESEE BUSCAR ", "MENU DE BUSQUEDA")
'If FECHA = "" Then Exit Sub

With DataEnvironment1
With .rsPARAMETRO_SEGUNDO_ORDEÑE
If .State = adStateOpen Then .Close
End With
Call .PARAMETRO_SEGUNDO_ORDEÑE(FECHA)
If .rsPARAMETRO_SEGUNDO_ORDEÑE.EOF Then
MsgBox "NO HAY NINGUN REGISTRO CON ESA FECHA DIGITE DE NUEVO, POR FAVOR DIGITE DE NUEVO"
Else
RPT_SEGUNDO_ORDEÑE.Show
End If
FRM_SEGUNDO_ORDEÑE.Hide
End With
End Sub


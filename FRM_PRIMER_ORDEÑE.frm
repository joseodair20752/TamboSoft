VERSION 5.00
Begin VB.Form FRM_PRIMER_ORDEÑE 
   Caption         =   "DATOS DEL PRIMER ORDEÑE"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   1515
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "DATOS DEL PRIMER ORDEÑE"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "FRM_PRIMER_ORDEÑE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim FECHA As Date
FECHA = InputBox(" INGRESE LA FECHA DEL ORDEÑE QUE DESEE BUSCAR ", "MENU DE BUSQUEDA")
'If FECHA = "" Then Exit Sub

With DataEnvironment1
With .rsPARAMETRO_PRIMER_ORDEÑE
If .State = adStateOpen Then .Close
End With
Call .PARAMETRO_PRIMER_ORDEÑE(FECHA)
If .rsPARAMETRO_PRIMER_ORDEÑE.EOF Then
MsgBox "NO HAY NINGUN REGISTRO CON ESA FECHA DIGITE DE NUEVO, POR FAVOR DIGITE DE NUEVO"
Else
RPTPRIMER_ORDEÑE.Show
End If
End With
FRM_PRIMER_ORDEÑE.Hide
End Sub

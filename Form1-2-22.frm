VERSION 5.00
Begin VB.Form Form1_EDAD_ANIMAL 
   Caption         =   "PARAMETROS EDAD DE ANIMALES"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "Form1-2-22.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESE EL RP DEL ANIMAL"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "EDAD, MESES, DÍAS"
      Height          =   3975
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      Begin VB.Image Image2 
         Height          =   2415
         Left            =   720
         Picture         =   "Form1-2-22.frx":0442
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2700
      End
   End
End
Attribute VB_Name = "Form1_EDAD_ANIMAL"
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
With .rsPARAMETRO_EDAD_ANIMAL
If .State = adStateOpen Then .Close
End With
Call .PARAMETRO_EDAD_ANIMAL(RP)
If .rsPARAMETRO_EDAD_ANIMAL.EOF Then
MsgBox "NO HAY NINGUN REGISTRO, POR FAVOR DIGITE DE NUEVO"
Else
DataReport1.Show
End If
End With

End Sub

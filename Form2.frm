VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "BUSQUEDA DE ANALISIS POR PARAMETROS"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1380
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "BUSQUEDA DE ANALISIS POR PARAMETROS"
      Height          =   855
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
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
With .rsPARAMETRO_ANALISIS
If .State = adStateOpen Then .Close
End With
Call .PARAMETRO_ANALISIS(RP)
If .rsPARAMETRO_ANALISIS.EOF Then
MsgBox "NO HAY NINGUN REGISTRO, POR FAVOR DIGITE DE NUEVO"
Else
RPT_PARAMETROS_ANALISIS.Show
End If
End With
Form2.Hide
End Sub

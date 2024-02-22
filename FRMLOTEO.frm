VERSION 5.00
Begin VB.Form FRMLOTEO 
   Caption         =   "LOTEO DE LOS ANIMALES DEL CAMPO"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form5"
   ScaleHeight     =   3375
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingreso de Loteo"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtCANTIDAD 
      DataField       =   "CANTIDAD"
      Height          =   285
      Left            =   1935
      TabIndex        =   9
      Top             =   1965
      Width           =   3375
   End
   Begin VB.TextBox txtRP 
      DataField       =   "RP"
      Height          =   285
      Left            =   1935
      TabIndex        =   7
      Top             =   1590
      Width           =   3375
   End
   Begin VB.TextBox txtCONDICIÓN 
      Height          =   285
      Left            =   1935
      TabIndex        =   5
      Top             =   1215
      Width           =   3375
   End
   Begin VB.TextBox txtRAZA 
      DataField       =   "RAZA"
      Height          =   285
      Left            =   1935
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtPOTRERO 
      DataField       =   "POTRERO"
      Height          =   285
      Left            =   1935
      TabIndex        =   1
      Top             =   450
      Width           =   3375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CANTIDAD:"
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   8
      Top             =   2010
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RP:"
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   6
      Top             =   1635
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CONDICIÓN:"
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RAZA:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   870
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "POTRERO:"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   1815
   End
End
Attribute VB_Name = "FRMLOTEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LOS DATOS ?", vbQuestion + vbYesNo, "CONFIRMACION")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR
With DataEnvironment1.rsLOTEO
.Open
.AddNew
 !POTRERO = txtPOTRERO.Text
 !RAZA = txtRAZA.Text
!CONDICIÓN = txtCONDICIÓN.Text
!RP = txtRP.Text
!CANTIDAD = txtCANTIDAD.Text

.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtPOTRERO = ""
txtRAZA = ""
txtCONDICIÓN = ""
txtRP = ""
txtCANTIDAD = ""

txtPOTRERO.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsLOTEO

.CancelUpdate
.Close
End With
End Sub

Private Sub Command2_Click()

frmeliminacion_loteo.Show
End Sub

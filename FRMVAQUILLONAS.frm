VERSION 5.00
Begin VB.Form FRMVAQUILLONAS 
   Caption         =   "INGRESO DE VAQUILLONAS"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6960
   LinkTopic       =   "Form4"
   ScaleHeight     =   4830
   ScaleWidth      =   6960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar Vaquillonas"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtRP_PADRE 
      DataField       =   "RP_PADRE"
      Height          =   285
      Left            =   2085
      TabIndex        =   9
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtRP_MADRE 
      DataField       =   "RP_MADRE"
      Height          =   285
      Left            =   2085
      TabIndex        =   7
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtPESO 
      DataField       =   "PESO"
      Height          =   285
      Left            =   2085
      TabIndex        =   5
      Top             =   1170
      Width           =   660
   End
   Begin VB.TextBox txtCANTIDAD_DE_DIENTES 
      DataField       =   "CANTIDAD_DE_DIENTES"
      Height          =   285
      Left            =   2085
      TabIndex        =   3
      Top             =   780
      Width           =   660
   End
   Begin VB.TextBox txtRP 
      DataField       =   "RP"
      Height          =   285
      Left            =   2085
      TabIndex        =   1
      Top             =   405
      Width           =   3375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RP_PADRE:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   1965
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RP_MADRE:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1590
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PESO:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1215
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CANTIDAD_DE_DIENTES:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   825
      Width           =   2055
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RP:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   450
      Width           =   1815
   End
End
Attribute VB_Name = "FRMVAQUILLONAS"
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
With DataEnvironment1.rsVAQUILLONAS
.Open
.AddNew
 !RP = txtRP.Text
 !CANTIDAD_DE_DIENTES = txtCANTIDAD_DE_DIENTES.Text
!PESO = txtPESO.Text
!RP_MADRE = txtRP_MADRE.Text
!RP_PADRE = txtRP_PADRE.Text
.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtRP = ""
txtCANTIDAD_DE_DIENTES = ""
txtPESO = ""
txtRP_MADRE = ""
txtRP_PADRE = ""
txtRP.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsVAQUILLONAS
.CancelUpdate
.Close
End With

End Sub

Private Sub Command2_Click()
FRMELIMINACIONVAQUILLONAS.Show
End Sub

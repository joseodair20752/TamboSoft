VERSION 5.00
Begin VB.Form FRMINGRESOCONTROLSOMATICO 
   Caption         =   "INGRESO DE CONTROL SOMATICO"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtVALOR_SEGUNDO_CONTROL 
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtVALOR_PRIMER_CONTROL 
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ELIMINAR CONTROLES SOMATICOS"
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESO DE CONTROL SOMATICO"
      Height          =   615
      Left            =   2040
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtSEGUNDO_CONTROL 
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   2625
      Width           =   1320
   End
   Begin VB.TextBox txtPRIMER_CONTROL 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   2250
      Width           =   1320
   End
   Begin VB.TextBox txtLOTE 
      Height          =   285
      Left            =   2955
      TabIndex        =   3
      Top             =   545
      Width           =   3375
   End
   Begin VB.TextBox txtRP 
      Height          =   285
      Left            =   2955
      TabIndex        =   1
      Top             =   165
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "VALOR_SEGUNDO_CONTROL"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "VALOR_PRIMER_CONTROL"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SEGUNDO_CONTROL:"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   2550
      Width           =   2295
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PRIMER_CONTROL:"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   2175
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "LOTE:"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RP:"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   210
      Width           =   1815
   End
End
Attribute VB_Name = "FRMINGRESOCONTROLSOMATICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿ CONFIRMA EL INGRESO DE LOS DATOS?", vbQuestion + vbYesNo, "CONFIRMACION")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR
With DataEnvironment1.rsCONTROL_SOMATICO
.Open
.AddNew
!RP = txtRP.Text
!VALOR_PRIMER_CONTROL = txtVALOR_PRIMER_CONTROL.Text
!LOTE = txtLOTE.Text
!VALOR_SEGUNDO_CONTROL = txtVALOR_SEGUNDO_CONTROL.Text
!PRIMER_CONTROL = txtPRIMER_CONTROL.Text
!SEGUNDO_CONTROL = txtSEGUNDO_CONTROL.Text
.Update
.Close
MsgBox " LOS DATOS FUERON ALMACENADOS CON ÉXITO", vbInformation, "INFORMACION"
txtRP = ""
txtVALOR_PRIMER_CONTROL = ""
txtLOTE = ""
txtVALOR_SEGUNDO_CONTROL = ""
txtPRIMER_CONTROL = ""
txtSEGUNDO_CONTROL = ""
txtRP.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCONTROL_SOMATICO
.CancelUpdate
.Close
End With
End Sub

Private Sub Command2_Click()
FRMCONTROLSOMATICO.Show
End Sub

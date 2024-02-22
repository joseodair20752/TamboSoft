VERSION 5.00
Begin VB.Form FRMSEGUNDORDEÑE 
   Caption         =   "INGRESO DEL SEGUNDO ORDEÑE"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCANTIDAD_EN_TANQUE 
      Height          =   285
      Left            =   3600
      TabIndex        =   14
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ELIMINAR"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESO DE 2° ORDEÑE"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtFIN 
      Height          =   285
      Left            =   3615
      TabIndex        =   11
      Top             =   2145
      Width           =   1320
   End
   Begin VB.TextBox txtHORA_DEL_SEGUNDO_ORDEÑE 
      Height          =   285
      Left            =   3615
      TabIndex        =   9
      Top             =   1755
      Width           =   1320
   End
   Begin VB.TextBox txtFECHA_DE_ENTREGA_AL_LABORATORIO 
      DataField       =   "FECHA_DE_ENTREGA_AL_LABORATORIO"
      Height          =   285
      Left            =   3615
      TabIndex        =   7
      Top             =   1380
      Width           =   1320
   End
   Begin VB.TextBox txtLABORATORIO 
      DataField       =   "LABORATORIO"
      Height          =   285
      Left            =   3615
      TabIndex        =   5
      Top             =   1005
      Width           =   3375
   End
   Begin VB.TextBox txtINSPECTOR 
      Height          =   285
      Left            =   3615
      TabIndex        =   3
      Top             =   615
      Width           =   3375
   End
   Begin VB.TextBox txtFECHA_CONTROL 
      Height          =   285
      Left            =   3615
      TabIndex        =   1
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "CANTIDAD_EN_TANQUE"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "FIN:"
      Height          =   255
      Index           =   5
      Left            =   1770
      TabIndex        =   10
      Top             =   2190
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "HORA_DEL_SEGUNDO_ORDEÑE:"
      Height          =   255
      Index           =   4
      Left            =   330
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "FECHA_DE_ENTREGA_AL_LABORATORIO:"
      Height          =   255
      Index           =   3
      Left            =   210
      TabIndex        =   6
      Top             =   1425
      Width           =   3375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "LABORATORIO:"
      Height          =   255
      Index           =   2
      Left            =   1770
      TabIndex        =   4
      Top             =   1050
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "INSPECTOR:"
      Height          =   255
      Index           =   1
      Left            =   1770
      TabIndex        =   2
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "FECHA_CONTROL:"
      Height          =   255
      Index           =   0
      Left            =   1770
      TabIndex        =   0
      Top             =   285
      Width           =   1815
   End
End
Attribute VB_Name = "FRMSEGUNDORDEÑE"
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
With DataEnvironment1.rsCARGA_DE_CONTROLES_DEL_SEGUNDO_ORDEÑE
.Open
.AddNew
    !LABORATORIO = txtLABORATORIO.Text
    !FECHA_CONTROL = txtFECHA_CONTROL.Text
    !INSPECTOR = txtINSPECTOR.Text
    !FECHA_DE_ENTREGA_AL_LABORATORIO = txtFECHA_DE_ENTREGA_AL_LABORATORIO.Text
    !HORA_DEL_SEGUNDO_ORDEÑE = txtHORA_DEL_SEGUNDO_ORDEÑE.Text
    !CANTIDAD_EN_TANQUE = txtCANTIDAD_EN_TANQUE.Text
    !FIN = txtFIN.Text

.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtLABORATORIO = ""
txtFECHA_CONTROL = ""
txtINSPECTOR = ""
txtFECHA_DE_ENTREGA_AL_LABORATORIO = ""
txtHORA_DEL_SEGUNDO_ORDEÑE = ""
txtFIN = ""
txtCANTIDAD_EN_TANQUE = ""
txtFECHA_CONTROL.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_CONTROLES_DEL_SEGUNDO_ORDEÑE
.CancelUpdate
.Close
End With
End Sub

Private Sub Command2_Click()
FrmEliminación_Segundo_Ordeñe.Show
End Sub

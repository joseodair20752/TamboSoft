VERSION 5.00
Begin VB.Form FRMORDEÑE1 
   Caption         =   "Ingresos_Analisis"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCANTIDAD_EN_TANQUE 
      Height          =   285
      Left            =   4320
      TabIndex        =   14
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ELIMINAR"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESO DEL PRIMER ORDEÑE"
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox txtFIN 
      Height          =   285
      Left            =   4380
      TabIndex        =   11
      Top             =   2155
      Width           =   1320
   End
   Begin VB.TextBox txtHORA_DEL_PRIMER_ORDEÑE 
      Height          =   285
      Left            =   4380
      TabIndex        =   9
      Top             =   1775
      Width           =   1320
   End
   Begin VB.TextBox txtFECHA_DE_ENTREGA_AL_LABORATORIO 
      Height          =   285
      Left            =   4380
      TabIndex        =   7
      Top             =   1395
      Width           =   1320
   End
   Begin VB.TextBox txtLABORATORIO 
      Height          =   285
      Left            =   4380
      TabIndex        =   5
      Top             =   1015
      Width           =   3375
   End
   Begin VB.TextBox txtINSPECTOR 
      Height          =   285
      Left            =   4380
      TabIndex        =   3
      Top             =   635
      Width           =   3375
   End
   Begin VB.TextBox txtFECHA_CONTROL 
      Height          =   285
      Left            =   4380
      TabIndex        =   1
      Top             =   255
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "CANTIDAD_EN_TANQUE"
      Height          =   255
      Left            =   1320
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
      Left            =   2520
      TabIndex        =   10
      Top             =   2205
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "HORA_DEL_PRIMER_ORDEÑE:"
      Height          =   255
      Index           =   4
      Left            =   1695
      TabIndex        =   8
      Top             =   1815
      Width           =   2655
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "FECHA_DE_ENTREGA_AL_LABORATORIO:"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "LABORATORIO:"
      Height          =   255
      Index           =   2
      Left            =   2535
      TabIndex        =   4
      Top             =   1065
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "INSPECTOR:"
      Height          =   255
      Index           =   1
      Left            =   2535
      TabIndex        =   2
      Top             =   675
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "FECHA_CONTROL:"
      Height          =   255
      Index           =   0
      Left            =   2535
      TabIndex        =   0
      Top             =   300
      Width           =   1815
   End
End
Attribute VB_Name = "FRMORDEÑE1"
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
With DataEnvironment1.rsCARGA_DE_CONTROLES_DEL_PRIMER_ORDEÑE
.Open
.AddNew
 !LABORATORIO = txtLABORATORIO.Text
 !FECHA_CONTROL = txtFECHA_CONTROL.Text
!INSPECTOR = txtINSPECTOR.Text
!FECHA_DE_ENTREGA_AL_LABORATORIO = txtFECHA_DE_ENTREGA_AL_LABORATORIO.Text
!HORA_DEL_PRIMER_ORDEÑE = txtHORA_DEL_PRIMER_ORDEÑE.Text
!FIN = txtFIN.Text
!CANTIDAD_EN_TANQUE = txtCANTIDAD_EN_TANQUE.Text
.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtLABORATORIO = ""
txtFECHA_CONTROL = ""
txtINSPECTOR = ""
txtFECHA_DE_ENTREGA_AL_LABORATORIO = ""
txtHORA_DEL_PRIMER_ORDEÑE = ""
txtFIN = ""
txtCANTIDAD_EN_TANQUE = ""
txtFECHA_CONTROL.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_CONTROLES_DEL_PRIMER_ORDEÑE
.CancelUpdate
.Close
End With

End Sub

Private Sub Command2_Click()
FRMELIMINACION_PRIMER_ORDEÑE.Show




End Sub

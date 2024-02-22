VERSION 5.00
Begin VB.Form FRMTOROS 
   Caption         =   "INGRESO DE TOROS"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "CANCELAR"
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton CMDGRABAR 
      Caption         =   "ACEPTAR"
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtFECHA_DE_NACIMIENTO 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtRP 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtNOMBRE 
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "FECHA DE NACIMIENTO "
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "RP"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "NOMBRE"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FRMTOROS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDGRABAR_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LOS DATOS ?", vbQuestion + vbYesNo, "CONFIRMACION ")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR

With DataEnvironment1.rsCARGA_DE_TOROS
.Open
.AddNew
!NOMBRE = txtNOMBRE.Text
!RP = txtRP.Text
!FECHA_DE_NACIMIENTO = txtFECHA_DE_NACIMIENTO.Text
.Update
.Close
MsgBox "DATOS ALMACENADOS CON ÉXITO", vbInformation, "INFORMACIÓN"

txtNOMBRE = ""
txtRP = ""
txtFECHA_DE_NACIMIENTO = ""
txtNOMBRE.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_TOROS
.CancelUpdate
.Close
End With
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
MsgBox (" INGRESE EL RP DEL ANIMAL POR FAVOR"), vbInformation
End Sub

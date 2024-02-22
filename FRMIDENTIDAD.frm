VERSION 5.00
Begin VB.Form FRMIDENTIDAD 
   Caption         =   "INGRESO DE LA IDENTIDAD"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "ELIMINAR"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESO"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtCANT_TAMBOS 
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   7
      Top             =   1320
      Width           =   660
   End
   Begin VB.TextBox txtCANT_PRODUCTORES 
      Height          =   285
      Left            =   1995
      TabIndex        =   5
      Top             =   940
      Width           =   660
   End
   Begin VB.TextBox txtNOMBRE 
      Height          =   285
      Left            =   1995
      TabIndex        =   3
      Top             =   560
      Width           =   3375
   End
   Begin VB.TextBox txtNUMERO 
      Height          =   285
      Left            =   1995
      TabIndex        =   1
      Top             =   180
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CANT_TAMBOS:"
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   6
      Top             =   1365
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CANT_PRODUCTORES:"
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   4
      Top             =   990
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "NOMBRE:"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "NUMERO:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   225
      Width           =   1815
   End
End
Attribute VB_Name = "FRMIDENTIDAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LOS DATOS ?", vbQuestion + vbYesNo, "CONFIRMACION ")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR

With DataEnvironment1.rsCARGA_DE_DATOS_ESTABLECIMIENTO
.Open
.AddNew
!NUMERO = txtNUMERO.Text
!NOMBRE = txtNOMBRE.Text
!CANT_PRODUCTORES = txtCANT_PRODUCTORES.Text
!CANT_TAMBOS = txtCANT_TAMBOS.Text
.Update
.Close
MsgBox "DATOS ALMACENADOS CON ÉXITO", vbInformation, "INFORMACIÓN"

txtNUMERO = ""
txtCANT_PRODUCTORES = ""
txtCANT_TAMBOS = ""
txtNUMERO.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_DATOS_ESTABLECIMIENTO
.CancelUpdate
.Close
End With
End Sub


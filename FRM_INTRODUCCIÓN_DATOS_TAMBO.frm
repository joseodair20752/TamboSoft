VERSION 5.00
Begin VB.Form FRM_INTRODUCCIÓN_DATOS_TAMBO 
   Caption         =   "INTRODUCCIÓN DE DATOS TAMBO"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtLocalidad 
      DataField       =   "Localidad"
      Height          =   285
      Left            =   2475
      TabIndex        =   9
      Top             =   1865
      Width           =   3375
   End
   Begin VB.TextBox txtDirección 
      DataField       =   "Dirección"
      Height          =   285
      Left            =   2475
      TabIndex        =   7
      Top             =   1485
      Width           =   3375
   End
   Begin VB.TextBox txtNumero_Establecimiento 
      DataField       =   "Numero_Establecimiento"
      Height          =   285
      Left            =   2475
      TabIndex        =   5
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtNumero 
      DataField       =   "Numero"
      Height          =   285
      Left            =   2475
      TabIndex        =   3
      Top             =   725
      Width           =   660
   End
   Begin VB.TextBox txtId 
      DataField       =   "Id"
      Height          =   285
      Left            =   2475
      TabIndex        =   1
      Top             =   345
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Localidad:"
      Height          =   255
      Index           =   4
      Left            =   630
      TabIndex        =   8
      Top             =   1905
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      Height          =   255
      Index           =   3
      Left            =   630
      TabIndex        =   6
      Top             =   1530
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Numero_Establecimiento:"
      Height          =   255
      Index           =   2
      Left            =   630
      TabIndex        =   4
      Top             =   1155
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Numero:"
      Height          =   255
      Index           =   1
      Left            =   630
      TabIndex        =   2
      Top             =   765
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Id:"
      Height          =   255
      Index           =   0
      Left            =   630
      TabIndex        =   0
      Top             =   390
      Width           =   1815
   End
End
Attribute VB_Name = "FRM_INTRODUCCIÓN_DATOS_TAMBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LOS DATOS DEL TAMBO?", vbQuestion + vbYesNo, "CONFIRMACION")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR
With DataEnvironment1.rsDatos_del_Tambo
.Open
.AddNew
!id = txtId.Text
!NUMERO = txtNumero.Text
!Numero_Establecimiento = txtNumero_Establecimiento.Text
!Dirección = txtDirección.Text
!Localidad = txtLocalidad.Text
.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtId = ""
txtNumero = ""
txtNumero_Establecimiento = ""
txtDirección = ""
txtLocalidad = ""
txtId.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsDatos_del_Tambo
.CancelUpdate
.Close
End With

End Sub

Private Sub Command2_Click()
FRM_INTRODUCCIÓN_DATOS_TAMBO.Hide
End Sub

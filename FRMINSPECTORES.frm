VERSION 5.00
Begin VB.Form FRMINSPECTORES 
   Caption         =   "INGRESO DE LOS INSPECTORES"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   3960
      TabIndex        =   15
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESO DE INSPECTORES"
      Height          =   615
      Left            =   720
      TabIndex        =   14
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox txtTELEFONO 
      Height          =   285
      Left            =   2190
      TabIndex        =   13
      Top             =   2670
      Width           =   3375
   End
   Begin VB.TextBox txtCODIGO_POSTAL 
      Height          =   285
      Left            =   2190
      TabIndex        =   11
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtLOCALIDAD 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1905
      Width           =   3375
   End
   Begin VB.TextBox txtDEPTO 
      Height          =   285
      Left            =   2190
      TabIndex        =   7
      Top             =   1530
      Width           =   3375
   End
   Begin VB.TextBox txtDIRECCION 
      Height          =   285
      Left            =   2190
      TabIndex        =   5
      Top             =   1155
      Width           =   3375
   End
   Begin VB.TextBox txtNOMBRE 
      Height          =   285
      Left            =   2190
      TabIndex        =   3
      Top             =   765
      Width           =   3375
   End
   Begin VB.TextBox txtAPELLIDO 
      Height          =   285
      Left            =   2190
      TabIndex        =   1
      Top             =   390
      Width           =   3375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "TELEFONO:"
      Height          =   255
      Index           =   6
      Left            =   345
      TabIndex        =   12
      Top             =   2715
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CODIGO_POSTAL:"
      Height          =   255
      Index           =   5
      Left            =   345
      TabIndex        =   10
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "LOCALIDAD:"
      Height          =   255
      Index           =   4
      Left            =   345
      TabIndex        =   8
      Top             =   1950
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DEPTO:"
      Height          =   255
      Index           =   3
      Left            =   345
      TabIndex        =   6
      Top             =   1575
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DIRECCION:"
      Height          =   255
      Index           =   2
      Left            =   345
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "NOMBRE:"
      Height          =   255
      Index           =   1
      Left            =   345
      TabIndex        =   2
      Top             =   810
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "APELLIDO:"
      Height          =   255
      Index           =   0
      Left            =   345
      TabIndex        =   0
      Top             =   435
      Width           =   1815
   End
End
Attribute VB_Name = "FRMINSPECTORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LOS DATOS DE ABORTO?", vbQuestion + vbYesNo, "CONFIRMACION")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR
With DataEnvironment1.rsALTAS_DE_INSPECTOR

.Open
.AddNew
 !APELLIDO = txtAPELLIDO.Text
 !NOMBRE = txtNOMBRE.Text
!DIRECCION = txtDIRECCION.Text
!DEPTO = txtDEPTO.Text
!CODIGO_POSTAL = txtCODIGO_POSTAL.Text
!TELEFONO = txtTELEFONO.Text
!Localidad = txtLOCALIDAD.Text
.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtAPELLIDO = ""
txtNOMBRE = ""
txtDIRECCION = ""
txtDEPTO = ""
txtCODIGO_POSTAL = ""
txtTELEFONO = ""
txtLOCALIDAD = ""
txtAPELLIDO.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsALTAS_DE_INSPECTOR
.Update
.Close
End With
End Sub

Private Sub Command2_Click()
FRMINSPECTORES.Hide
FRMELIMINACIONINSPECTORES.Show

End Sub

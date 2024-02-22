VERSION 5.00
Begin VB.Form frmHabilitacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu Habilitar"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7515
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdIngresar 
      Caption         =   "Ingresar"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtContraseña 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Text            =   "AGENDA MEDICA"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave de habilitación"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bienbenido a la pantalla de Habilitación del programa"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "frmHabilitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdIngresar_Click()
If txtUsuario = "" Then
MsgBox " Debe de ingresar un nombre de usuario ", vbInformation, "Información"
txtUsuario.SetFocus
Exit Sub
End If

If txtContraseña = "" Then
MsgBox " Debe de ingresar una contraseña", vbInformation, "Información"
txtContraseña.SetFocus
Exit Sub
End If

If Trim(UCase(txtUsuario)) = "Usted ha ingresado a su agenda " And Trim(UCase(txtContraseña)) = "212358020014" Then
MsgBox " Usted ha ingresado al programa Software de tambo 1.0", vbExclamation, "INGRESO"
frmBase.Show
frmHabilitacion.Hide
Else
MsgBox " No trate de entrar sin permiso luego de tres intentos fallidos el programa quedará nulo por unos días", vbCritical, "ILEGAL"
txtUsuario.SetFocus
txtUsuario.SelStart = 0
txtUsuario.SelLength = Len(txtUsuario)


End If

End Sub


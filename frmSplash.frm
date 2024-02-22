VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4395
   ClientLeft      =   2430
   ClientTop       =   2490
   ClientWidth     =   6615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SOFTWARE DE TAMBO 1.0"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   1875
      TabIndex        =   0
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   4575
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6660
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Animation1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  frmSplash.Hide
frmHabilitacion.Show
  
 
End Sub

Private Sub Form_Load()
'MsgBox "Esto es una demostración del programa SOFTWARE DE TAMBO 1.0"
End Sub

Private Sub Frame1_Click()
    'Unload Me
End Sub

Private Sub Timer1_Timer()
'Timer1.Interval = 1000
' frmSplash.Hide
' frmBase.Show
End Sub

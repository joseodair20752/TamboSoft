VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tempo 
      Left            =   1200
      Top             =   720
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtClave 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblIntentos 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblTiempo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblClave 
      Caption         =   "Contraseña:"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre:"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim bdUsuarios As Database, rstUsuarios As Recordset, strRuta As String
Dim I As Integer, Intentos As Integer, strNombre As String
Dim strUser As String, strPass As String, Acceso As Boolean, Admin As Boolean

'Private Sub cmdAceptar_Click()
'    If Trim(txtNombre.Text) = "" Then
'        MsgBox "Debes escribir un Nombre de Usuario para continuar"
'        Exit Sub
'    ElseIf Trim(txtClave.Text) = "" Then
'        MsgBox "Debes de Escribir una Clave", vbCritical, "Error"
'        Exit Sub
'    Else
'        Intentos = Intentos - 1
'        lblIntentos.Caption = Intentos
'        If Intentos <= 0 Then
'            MsgBox "Agotó los intentos permitidos" & _
'            vbCrLf & "Consulte con el Administrador del sistema" _
'            , vbInformation, "Finalizado"
'            End
'        End If
'    End If
'    With rstUsuarios
'    If (.BOF) = True And (.EOF) = True Then
'        MsgBox " la base de datos está Vacia", vbCritical, "BAse de Datos"
'        Exit Sub
'    End If
'    End With
'    Acceso = Validar(Trim(txtNombre), Trim(txtClave))
'    If Acceso = True Then
'        MsgBox "Se ha validado el Usuario", vbExclamation, "Validación"
'    Else
'    MsgBox "Usuario o clave errada", vbCritical, "Validación"
'    End If
'End Sub

Private Sub cmdCancelar_Click()
    End
End Sub

'Private Sub Form_Load()
'    I = 30: Intentos = 5
'    strRuta = App.Path
'    Set bdUsuarios = OpenDatabase(strRuta & "\basedat.mdb")
'    Set rstUsuarios = bdUsuarios.OpenRecordset("Usuarios")
'    Tempo.Interval = 1000
'    Tempo.Enabled = True
'End Sub
'
'Private Sub Form_Paint()
'    lblIntentos.Caption = Intentos
'End Sub
'
'Private Sub Tempo_Timer()
'   lblTiempo.Caption = I
'   I = I - 1
'   If I < 0 Then
'       MsgBox "Tiempo Finalizado", vbInformation, "Tiempo"
'       End
'   End If
'End Sub
'
'Private Function Validar(strNombre As String, strClave As String) As Boolean
'    With rstUsuarios
'        If (.BOF) = True And (.EOF) = True Then
'            Exit Function
'        End If
'        .Index = "IndSeudom"
'        .Seek "=", Trim(strNombre)
'        If .NoMatch Then
'            Exit Function
'        End If
'        If Trim(!Password) = Trim(strClave) Then
'            Validar = True
'            strNombre = !NOMBRE
'            Admin = !Admin
'        End If
'    End With
'End Function


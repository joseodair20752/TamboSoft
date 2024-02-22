VERSION 5.00
Begin VB.Form frm_Conaprole 
   Caption         =   "Ingreso a la Pagina de Internet de Conaprole"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form3"
   ScaleHeight     =   4515
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VER PAGINA DE CONAPROLE EN LA RED"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Width           =   4215
   End
End
Attribute VB_Name = "frm_Conaprole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Public Explorer As SHDocVw.InternetExplorer


'Private Sub Command1_Click()
'
'
' On Error GoTo manejadorerror
'    Set Explorer = New SHDocVw.InternetExplorer
'    Explorer.Visible = True
'    Explorer.navigate Combo1.Text
'    Exit Sub
'manejadorerror:
'    MsgBox "Error visualizando el archivo", , Err.Description
'End Sub

'Private Sub Form_Load()
''Añadir unos pocos servidores web al cuadro combo durante
''la puesta en marcha
'    'página inicial de Microsoft Corp.
'    Combo1.AddItem "http://www.conaprole.com/"
'' Combo1.AddItem "http://www.windx.com"
'' Combo1.AddItem "http://www.microsoft.com/vbasic/"
'    'página inicial de Microsoft Press
'    'Combo1.AddItem "http://mspress.microsoft.com/"
'    'página inicial de Microsoft Visual Basic Programming
'    'Combo1.AddItem "http://www.microsoft.com/vbasic/"
'    'recursos de Fawcette Publication para programación en VB
'   ' Combo1.AddItem "http://www.windx.com"
'    'página inicial de VB de Carl y Gary (no-Microsoft)
'   ' Combo1.AddItem "http://www.apexsc.com/vb/"
'End Sub
'
'

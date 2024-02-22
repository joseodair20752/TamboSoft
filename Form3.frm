VERSION 5.00
Begin VB.Form FRMZOMATICO 
   Caption         =   "CONTROL ZOMATICO"
   ClientHeight    =   4545
   ClientLeft      =   3360
   ClientTop       =   2985
   ClientWidth     =   7110
   LinkTopic       =   "Form3"
   ScaleHeight     =   4545
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "VER LISTADO  DE EVOLUCION SOMATICA"
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   3120
      Width           =   2295
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Chart.8"
      DisplayType     =   1  'Icon
      Height          =   1335
      Left            =   1920
      OleObjectBlob   =   "Form3.frx":0000
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.Line Line4 
      X1              =   5760
      X2              =   1080
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line3 
      X1              =   1080
      X2              =   1080
      Y1              =   360
      Y2              =   4200
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   1080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   5760
      Y1              =   360
      Y2              =   4200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Haga doble click en la figura presente en este formulario"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2280
      Width           =   3255
   End
End
Attribute VB_Name = "FRMZOMATICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RPTLISTA_GENERAL_DE_CONTROL_ZOMATICO.Show
End Sub
'SELECT Grafico_Consulta_Somatico.`VALOR_PRIMER_CONTROL`, Grafico_Consulta_Somatico.`VALOR_SEGUNDO_CONTROL` From `Grafico_Consulta_Somatico` Grafico_Consulta_Somatico Order By Grafico_Consulta_Somatico.`VALOR_PRIMER_CONTROL` ASC, Grafico_Consulta_Somatico.`VALOR_SEGUNDO_CONTROL` ASC
Private Sub Form_Load()

End Sub

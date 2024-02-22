VERSION 5.00
Begin VB.Form FRMTANQUE_GRAFICO 
   Caption         =   "GRAFICO DE CANTIDAD EN TANQUE EN PRIMER ORDEï¿½E"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form5"
   ScaleHeight     =   2235
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      DisplayType     =   1  'Icon
      Height          =   615
      Left            =   2160
      OleObjectBlob   =   "FRMTANQUE_GRAFICO.frx":0000
      SourceDoc       =   "C:\MIS DOCUMENTOS\SOFTWARE_DE_TAMBO1.0\CONTROLES_PRIMER_ORDEÃ‘E.XLS"
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HAGA DOBLE CLICK EN LA IMAGEN SUPERIOR PARA VER EL GRAFICO DEL PRIMER ORDEÑE."
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
   End
End
Attribute VB_Name = "FRMTANQUE_GRAFICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()


FRMTANQUE_GRAFICO.Hide

End Sub

Private Sub Command2_Click()
OLE1.Action = 1
End Sub


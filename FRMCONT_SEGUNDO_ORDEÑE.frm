VERSION 5.00
Begin VB.Form FRMCONT_SEGUNDO_ORDEÑE 
   Caption         =   "GRAFICO DEL SEGUNDO ORDEÑE"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HAGA DOBLE CLICK EN LA IMGEN SUPERIOR PARA SABER EL CONTROL DEL SEGUNDO ORDEÑE"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   4095
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   1440
      OleObjectBlob   =   "FRMCONT_SEGUNDO_ORDEÑE.frx":0000
      SourceDoc       =   "C:\MIS DOCUMENTOS\SOFTWARE_DE_TAMBO1.0\CONTROLES SEGUNDO ORDEÑE.XLS"
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FRMCONT_SEGUNDO_ORDEÑE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

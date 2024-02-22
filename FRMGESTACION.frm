VERSION 5.00
Begin VB.Form FRMGESTACION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MESES DE GESTACION"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   1560
      OleObjectBlob   =   "FRMGESTACION.frx":0000
      SourceDoc       =   "C:\Users\JOSE CAMACHO\Desktop\Sofware_MisPRUEBAS\MESES DE EMBARAZO.xls"
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Haga doble click en la figura presente en este formulario "
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FRMGESTACION.frx":1C18
      Height          =   1095
      Left            =   9120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "FRMGESTACION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

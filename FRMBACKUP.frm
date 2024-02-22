VERSION 5.00
Begin VB.Form FRMBACKUP 
   Caption         =   "FORMULARIO DE RESPALDO BASE TAMBO"
   ClientHeight    =   9375
   ClientLeft      =   11895
   ClientTop       =   7575
   ClientWidth     =   16350
   LinkTopic       =   "Form4"
   ScaleHeight     =   9375
   ScaleWidth      =   16350
   Begin VB.OLE OLE1 
      Class           =   "Excel.SheetBinaryMacroEnabled.12"
      Height          =   6495
      Left            =   450
      OleObjectBlob   =   "FRMBACKUP.frx":0000
      TabIndex        =   1
      Top             =   330
      Width           =   14775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Haga doble click en la figura superior para efectuar su copia de respaldo"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   7560
      Width           =   3855
   End
End
Attribute VB_Name = "FRMBACKUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub


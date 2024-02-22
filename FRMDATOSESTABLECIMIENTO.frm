VERSION 5.00
Begin VB.Form FRMDATOSESTABLECIMIENTO 
   Caption         =   "DATOS ESTABLECIMIENTO"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtLOCALIDAD 
      DataField       =   "LOCALIDAD"
      DataMember      =   "CARGA_DE_DATOS_ESTABLECIMIENTO"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2460
      TabIndex        =   9
      Top             =   1700
      Width           =   3375
   End
   Begin VB.TextBox txtDIRECCION 
      DataField       =   "DIRECCION"
      DataMember      =   "CARGA_DE_DATOS_ESTABLECIMIENTO"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2460
      TabIndex        =   7
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtNOMBRE_PRODUCTOR 
      DataField       =   "NOMBRE_PRODUCTOR"
      DataMember      =   "CARGA_DE_DATOS_ESTABLECIMIENTO"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2460
      TabIndex        =   5
      Top             =   940
      Width           =   3375
   End
   Begin VB.TextBox txtNOMBRE_ESTABLECIMIENTO 
      DataField       =   "NOMBRE_ESTABLECIMIENTO"
      DataMember      =   "CARGA_DE_DATOS_ESTABLECIMIENTO"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2460
      TabIndex        =   3
      Top             =   560
      Width           =   3375
   End
   Begin VB.TextBox txtNUMERO 
      DataField       =   "NUMERO"
      DataMember      =   "CARGA_DE_DATOS_ESTABLECIMIENTO"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2460
      TabIndex        =   1
      Top             =   180
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "LOCALIDAD:"
      Height          =   255
      Index           =   4
      Left            =   615
      TabIndex        =   8
      Top             =   1740
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DIRECCION:"
      Height          =   255
      Index           =   3
      Left            =   615
      TabIndex        =   6
      Top             =   1365
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "NOMBRE_PRODUCTOR:"
      Height          =   255
      Index           =   2
      Left            =   615
      TabIndex        =   4
      Top             =   990
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "NOMBRE_ESTABLECIMIENTO:"
      Height          =   255
      Index           =   1
      Left            =   615
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "NUMERO:"
      Height          =   255
      Index           =   0
      Left            =   615
      TabIndex        =   0
      Top             =   225
      Width           =   1815
   End
End
Attribute VB_Name = "FRMDATOSESTABLECIMIENTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FRMDATOSESTABLECIMIENTO.Hide
End Sub

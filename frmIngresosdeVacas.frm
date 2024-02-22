VERSION 5.00
Begin VB.Form frmIngresosdeVacas 
   Caption         =   "INGRESO DE VACAS"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   15195
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFECHA_NACIMIENTO 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   72
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   7320
      TabIndex        =   67
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4080
      TabIndex        =   66
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox txt5Lac 
      Height          =   285
      Left            =   5160
      TabIndex        =   28
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtGr5 
      Height          =   285
      Left            =   8160
      TabIndex        =   27
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtProt6 
      Height          =   285
      Left            =   10200
      TabIndex        =   26
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox txt4Lac 
      Height          =   285
      Left            =   5160
      TabIndex        =   25
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtGr6 
      Height          =   285
      Left            =   8160
      TabIndex        =   24
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txt6Lac 
      Height          =   285
      Left            =   5160
      TabIndex        =   23
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtProt5 
      Height          =   285
      Left            =   10200
      TabIndex        =   22
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtProt4 
      Height          =   285
      Left            =   10200
      TabIndex        =   21
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtGr4 
      Height          =   285
      Left            =   8160
      TabIndex        =   20
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txt1Lac 
      Height          =   285
      Left            =   5160
      TabIndex        =   19
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txt2Lac 
      Height          =   285
      Left            =   5160
      TabIndex        =   18
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txt3Lac 
      Height          =   285
      Left            =   5160
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtGr1 
      Height          =   285
      Left            =   8160
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtGr2 
      Height          =   285
      Left            =   8160
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtGr3 
      Height          =   285
      Left            =   8160
      TabIndex        =   14
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtProt3 
      Height          =   285
      Left            =   10200
      TabIndex        =   13
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtProt1 
      Height          =   285
      Left            =   10200
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtProt2 
      Height          =   285
      Left            =   10200
      TabIndex        =   11
      Top             =   3360
      Width           =   735
   End
   Begin VB.ComboBox ComboPROCEDENCIA 
      Height          =   315
      ItemData        =   "frmIngresosdeVacas.frx":0000
      Left            =   240
      List            =   "frmIngresosdeVacas.frx":0013
      TabIndex        =   8
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "PADRE"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   2775
      Begin VB.TextBox txtNOMBRE_PADRE 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label39 
         Caption         =   "RP"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "MADRE"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2775
      Begin VB.TextBox txtNOMBRE_MADRE 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "RP"
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.ComboBox ComboRAZA 
      Height          =   315
      ItemData        =   "frmIngresosdeVacas.frx":0048
      Left            =   1440
      List            =   "frmIngresosdeVacas.frx":0055
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtMeses 
      Height          =   285
      Left            =   7320
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtAÑOS 
      Height          =   285
      Left            =   5160
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtRP 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "PRODUCCION"
      Height          =   5175
      Left            =   3240
      TabIndex        =   9
      Top             =   1440
      Width           =   8535
      Begin VB.TextBox txtCANT_DE_PARTOS 
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Text            =   "0"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label38 
         Caption         =   "Kg"
         Height          =   255
         Left            =   7800
         TabIndex        =   65
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label37 
         Caption         =   "Kg"
         Height          =   255
         Left            =   7800
         TabIndex        =   64
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label36 
         Caption         =   "Kg"
         Height          =   255
         Left            =   7800
         TabIndex        =   63
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label35 
         Caption         =   "Kg"
         Height          =   255
         Left            =   7800
         TabIndex        =   62
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "Kg"
         Height          =   255
         Left            =   7800
         TabIndex        =   61
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label33 
         Caption         =   "kg"
         Height          =   255
         Left            =   7800
         TabIndex        =   60
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label32 
         Caption         =   "Kg Prot"
         Height          =   255
         Left            =   6240
         TabIndex        =   59
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label31 
         Caption         =   "Kg Prot"
         Height          =   255
         Left            =   6240
         TabIndex        =   58
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label30 
         Caption         =   "Kg Prot"
         Height          =   255
         Left            =   6240
         TabIndex        =   57
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label29 
         Caption         =   "Kg Prot"
         Height          =   375
         Left            =   6240
         TabIndex        =   56
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label28 
         Caption         =   "Kg Prot"
         Height          =   375
         Left            =   6240
         TabIndex        =   55
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "Kg Prot"
         Height          =   255
         Left            =   6240
         TabIndex        =   54
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "al día de la Fecha"
         Height          =   255
         Left            =   6000
         TabIndex        =   53
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "Cantidad de Partos "
         Height          =   255
         Left            =   2520
         TabIndex        =   52
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "Lts       Grs"
         Height          =   255
         Left            =   3720
         TabIndex        =   51
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "Lts       Grs"
         Height          =   255
         Left            =   3720
         TabIndex        =   50
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Lts       Grs"
         Height          =   255
         Left            =   3720
         TabIndex        =   49
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Lts       Grs"
         Height          =   255
         Left            =   3720
         TabIndex        =   48
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Lts       Grs"
         Height          =   255
         Left            =   3720
         TabIndex        =   47
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Lts      Grs"
         Height          =   375
         Left            =   3720
         TabIndex        =   46
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Lec"
         Height          =   255
         Left            =   1320
         TabIndex        =   45
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Lec"
         Height          =   255
         Left            =   1320
         TabIndex        =   44
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Lec"
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Lec"
         Height          =   255
         Left            =   1320
         TabIndex        =   42
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Lec"
         Height          =   255
         Left            =   1320
         TabIndex        =   41
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Lec"
         Height          =   255
         Left            =   1320
         TabIndex        =   40
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "6 Lac"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "5 Lac"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "4 Lac"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "3 Lac"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "2 Lac"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "1 Lac"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "FECHA_NACIMIENTO:"
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label41 
      Caption         =   "Meses"
      Height          =   255
      Left            =   8280
      TabIndex        =   70
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Raza"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Años y "
      Height          =   255
      Left            =   6480
      TabIndex        =   32
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Edad"
      Height          =   255
      Left            =   4440
      TabIndex        =   31
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "PROCEDENCIA"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "RP"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmIngresosdeVacas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Command2_Click()
Dim RESPUESTA As Integer

RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LOS DATOS DE LOS ANIMALES?", vbQuestion + vbYesNo, "CONFIRMACIÓN")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR

With DataEnvironment1.rsCARGA_DE_VACAS
.Open
.AddNew

 !RP = txtRP.Text
!FECHA_NACIMIENTO = txtFECHA_NACIMIENTO.Text
!RAZA = ComboRAZA.Text
!NOMBRE_MADRE = txtNOMBRE_MADRE.Text
!NOMBRE_PADRE = txtNOMBRE_PADRE.Text
!PROCEDENCIA = ComboPROCEDENCIA.Text
!CANT_DE_PARTOS = txtCANT_DE_PARTOS.Text
!AÑOS = txtAÑOS.Text
!MESES = txtMeses.Text
If txtCANT_DE_PARTOS.Text = (0) Then
 txt1Lac.Visible = True
txt2Lac.Visible = True
 txt3Lac.Visible = True
 txt4Lac.Visible = True
txt5Lac.Visible = True
 txt6Lac.Visible = True
 txtGr1.Visible = True
 txtGr2.Visible = True
 txtGr3.Visible = True
 txtGr4.Visible = True
txtGr5.Visible = True
 txtGr6.Visible = True
 txtProt1.Visible = True
 txtProt2.Visible = True
 txtProt3.Visible = True
 txtProt4.Visible = True
 txtProt5.Visible = True
 txtProt6.Visible = True
!LACT1 = txt1Lac.Text = 0
!Gr1 = txtGr1.Text = 0
!Prot1 = txtProt1.Text = 0
!LACT2 = txt2Lac.Text = 0
!Gr2 = txtGr2.Text = 0
!Prot2 = txtProt2.Text = 0
!LACT3 = txt3Lac.Text = 0
!Gr3 = txtGr3.Text = 0
!Prot3 = txtProt3.Text = 0

!LACT4 = txt4Lac.Text = 0
!Gr4 = txtGr4.Text = 0
!Prot4 = txtProt4.Text = 0
!LACT5 = txt5Lac.Text = 0
!Gr5 = txtGr5.Text = 0
!Prot5 = txtProt5.Text = 0
!LACT6 = txt6Lac.Text = 0
!Gr6 = txtGr6.Text = 0
!Prot6 = txtProt6.Text = 0
End If



If txtCANT_DE_PARTOS.Text = 1 Then
!LACT1 = txt1Lac.Visible = True
!LACT2 = txt2Lac.Visible = False
!LACT3 = txt3Lac.Visible = False
!LACT4 = txt4Lac.Visible = False
!LACT5 = txt5Lac.Visible = False
!LACT6 = txt6Lac.Visible = False
!Gr1 = txtGr1.Visible = True
!Gr2 = txtGr2.Visible = False
!Gr3 = txtGr3.Visible = False
!Gr4 = txtGr4.Visible = False
!Gr5 = txtGr5.Visible = False
!Gr6 = txtGr6.Visible = False
!Prot1 = txtProt1.Visible = True
!Prot2 = txtProt2.Visible = False
!Prot3 = txtProt3.Visible = False
!Prot4 = txtProt4.Visible = False
!Prot5 = txtProt5.Visible = False
!Prot6 = txtProt6.Visible = False

!LACT1 = txt1Lac.Text
!Gr1 = txtGr1.Text
!Prot1 = txtProt1.Text
End If
If txtCANT_DE_PARTOS.Text = (2) Then
!LACT1 = txt1Lac.Visible = True
!LACT2 = txt2Lac.Visible = True
!LACT3 = txt3Lac.Visible = False
!LACT4 = txt4Lac.Visible = False
!LACT5 = txt5Lac.Visible = False
!LACT6 = txt6Lac.Visible = False
!Gr1 = txtGr1.Visible = True
!Gr2 = txtGr2.Visible = True
!Gr3 = txtGr3.Visible = False
!Gr4 = txtGr4.Visible = False
!Gr5 = txtGr5.Visible = False
!Gr6 = txtGr6.Visible = False
!Prot1 = txtProt1.Visible = True
!Prot2 = txtProt2.Visible = True
!Prot3 = txtProt3.Visible = False
!Prot4 = txtProt4.Visible = False
!Prot5 = txtProt5.Visible = False
!Prot6 = txtProt6.Visible = False

'!LACT2 = txt2Lac.Text
'!Gr2 = txtGr2.Text
'!Prot2 = txtProt2.Text
End If
If txtCANT_DE_PARTOS.Text = (3) Then
 txt1Lac.Visible = True
txt2Lac.Visible = True
 txt3Lac.Visible = True
 txt4Lac.Visible = False
 txt5Lac.Visible = False
 txt6Lac.Visible = False
 txtGr1.Visible = True
 txtGr2.Visible = True
 txtGr3.Visible = True
 txtGr4.Visible = False
 txtGr5.Visible = False
 txtGr6.Visible = False
 txtProt1.Visible = True
txtProt2.Visible = True
 txtProt3.Visible = True
 txtProt4.Visible = False
 txtProt5.Visible = False
 txtProt6.Visible = False

'!LACT3 = txt3Lac.Text
'!Gr3 = txtGr3.Text
'!Prot3 = txtProt3.Text

End If
If txtCANT_DE_PARTOS.Text = (4) Then
 txt1Lac.Visible = True
 txt2Lac.Visible = True
 txt3Lac.Visible = True
 txt4Lac.Visible = True
 txt5Lac.Visible = False
 txt6Lac.Visible = False
 txtGr1.Visible = True
 txtGr2.Visible = True
 txtGr3.Visible = True
 txtGr4.Visible = True
 txtGr5.Visible = False
 txtGr6.Visible = False
 txtProt1.Visible = True
 txtProt2.Visible = True
 txtProt3.Visible = True
 txtProt4.Visible = True
 txtProt5.Visible = False
 txtProt6.Visible = False

'!LACT4 = txt4Lac.Text
'!Gr4 = txtGr4.Text
'!Prot4 = txtProt4.Text
End If

If txtCANT_DE_PARTOS.Text = (5) Then
 txt1Lac.Visible = True
 txt2Lac.Visible = True
 txt3Lac.Visible = True
 txt4Lac.Visible = True
 txt5Lac.Visible = True
 txt6Lac.Visible = False
 txtGr1.Visible = True
 txtGr2.Visible = True
 txtGr3.Visible = True
 txtGr4.Visible = True
 txtGr5.Visible = True
 txtGr6.Visible = False
 txtProt1.Visible = True
 txtProt2.Visible = True
 txtProt3.Visible = True
 txtProt4.Visible = True
 txtProt5.Visible = True
 txtProt6.Visible = False

'!LACT5 = txt5Lac.Text
'!Gr5 = txtGr5.Text
'!Prot5 = txtProt5.Text
End If



If !CANT_DE_PARTOS = txtCANT_DE_PARTOS.Text = (6) Then
 txt1Lac.Visible = True
 txt2Lac.Visible = True
 txt3Lac.Visible = True
txt4Lac.Visible = True
txt5Lac.Visible = True
 txt6Lac.Visible = True
 txtGr1.Visible = True
 txtGr2.Visible = True
 txtGr3.Visible = True
 txtGr4.Visible = True
 txtGr5.Visible = True
 txtGr6.Visible = True
 txtProt1.Visible = True
 txtProt2.Visible = True
 txtProt3.Visible = True
 txtProt4.Visible = True
 txtProt5.Visible = True
 txtProt6.Visible = True

'!LACT6 = txt6Lac.Text
'!Gr6 = txtGr6.Text
'!Prot6 = txtProt6.Text

End If










'!LACT1 = txt1Lac.Text
'!LACT2 = txt2Lac.Text
'!LACT3 = txt3Lac.Text
'!LACT4 = txt4Lac.Text
'!LACT5 = txt5Lac.Text
''!LACT6 = txt6Lac.Text
'!Gr1 = txtGr1.Text
'!Gr2 = txtGr2.Text
'!Gr3 = txtGr3.Text
'!Gr4 = txtGr4.Text
'!Gr5 = txtGr5.Text
'!Gr6 = txtGr6.Text
'!Prot1 = txtProt1.Text
'!Prot2 = txtProt2.Text
'!Prot3 = txtProt3.Text
'!Prot4 = txtProt4.Text
'!Prot5 = txtProt5.Text
'!Prot6 = txtProt6.Text

.Update
.Close
MsgBox " Los datos de los animales fueron guardados con éxito", vbInformation, "Información"
txtRP = ""
txtFECHA_NACIMIENTO = ""
ComboRAZA = ""
txtNOMBRE_MADRE = ""
txtNOMBRE_PADRE = ""
ComboPROCEDENCIA = ""
txtAÑOS = ""
txtMeses = ""
txtCANT_DE_PARTOS = ""
txt1Lac = ""
txt2Lac = ""
txt3Lac = ""
txt4Lac = ""
txt5Lac = ""
txt6Lac = ""
txtGr1 = ""
txtGr2 = ""
txtGr3 = ""
txtGr4 = ""
txtGr5 = ""
txtGr6 = ""
txtProt1 = ""
txtProt2 = ""
txtProt3 = ""
txtProt4 = ""
txtProt5 = ""
txtProt6 = ""
txtRP.SetFocus


Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_VACAS
.CancelUpdate
.Close
End With

'If txtCANT_DE_PARTOS = (0) Then
'
'
'txt1Lac.Text = 0
' txt2Lac.Text = 0
'  txt3Lac.Text = 0
'  txt4Lac.Text = 0
'  txt5Lac.Text = 0
' txt6Lac.Text = 0
'
'txtGr1.Text = 0
' txtGr2.Text = 0
' txtGr3.Text = 0
' txtGr4.Text = 0
' txtGr5.Text = 0
'txtGr6.Text = 0
'
' txtProt1.Text = 0
' txtProt2.Text = 0
' txtProt3.Text = 0
' txtProt4.Text = 0
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'
'If txtCANT_DE_PARTOS = (1) Then
'txt1Lac.Text = txt1Lac.Text
' txt2Lac.Text = 0
'  txt3Lac.Text = 0
'  txt4Lac.Text = 0
'  txt5Lac.Text = 0
' txt6Lac.Text = 0
'
'txtGr1.Text = txtGr1.Text
' txtGr2.Text = 0
' txtGr3.Text = 0
' txtGr4.Text = 0
' txtGr5.Text = 0
'txtGr6.Text = 0
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = 0
' txtProt3.Text = 0
' txtProt4.Text = 0
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 2 Then
'txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
' txt3Lac.Text = 0
' txt4Lac.Text = 0
' txt5Lac.Text = 0
' txt6Lac.Text = 0
'
' txtGr1.Text = txtGr1.Text
' txtGr2.Text = txtGr2.Text
' txtGr3.Text = 0
' txtGr4.Text = 0
' txtGr5.Text = 0
' txtGr6.Text = 0
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = 0
' txtProt4.Text = 0
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 3 Then
'txt3Lac.Text = txt3Lac.Text
' txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
' txt4Lac.Text = 0
' txt5Lac.Text = 0
' txt6Lac.Text = 0
'
' txtGr1.Text = txtGr1.Text
' txtGr2.Text = txtGr2.Text
' txtGr3.Text = txtGr3.Text
' txtGr4.Text = 0
' txtGr5.Text = 0
' txtGr6.Text = 0
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = txtProt3.Text
' txtProt4.Text = 0
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 4 Then
' txt4Lac.Text = txt4Lac.Text
' txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
' txt4Lac.Text = txt3Lac.Text
' txt5Lac.Text = 0
' txt6Lac.Text = 0
'
'txtGr1.Text = txtGr1.Text
'txtGr2.Text = txtGr2.Text
' txtGr3.Text = txtGr3.Text
' txtGr4.Text = txtGr4.Text
' txtGr5.Text = 0
' txtGr6.Text = 0
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = txtProt3.Text
' txtProt4.Text = txtProt4.Text
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 5 Then
'txt5Lac.Text = txt5Lac.Text
' txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
'txt4Lac.Text = txt3Lac.Text
' txt4Lac.Text = txt4Lac.Text
' txt6Lac.Text = 0
'txtGr1.Text = txtGr1.Text
' txtGr2.Text = txtGr2.Text
' txtGr3.Text = txtGr3.Text
' txtGr4.Text = txtGr4.Text
' txtGr5.Text = txtGr5.Text
' txtGr6.Text = 0
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = txtProt3.Text
' txtProt4.Text = txtProt4.Text
' txtProt5.Text = txtProt5.Text
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 6 Then
'txt6Lac.Text = txt6Lac.Text
' txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
'txt3Lac.Text = txt3Lac.Text
' txt4Lac.Text = txt4Lac.Text
' txt5Lac.Text = txt5Lac.Text
'
'
' txtGr1.Text = txtGr1.Text
' txtGr2.Text = txtGr2.Text
' txtGr3.Text = txtGr3.Text
' txtGr4.Text = txtGr4.Text
' txtGr5.Text = txtGr5.Text
' txtGr6.Text = txtGr6.Text
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = txtProt3.Text
' txtProt4.Text = txtProt4.Text
' txtProt5.Text = txtProt5.Text
' txtProt6.Text = txtProt6.Text
'End If
'
'
'
End With
End If
End Sub
Private Sub Command3_Click()
Exit Sub
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
'MonthView1.Day = 1
End Sub

Private Sub txtCANT_DE_PARTOS_Change()
'If txtCANT_DE_PARTOS.Text = "" Or (0) Then
'Exit Sub
'
'If txtCANT_DE_PARTOS = (1) Then
'
'
'txt1Lac.Visible = True
'txt2Lac.Visible = False
'txt3Lac.Visible = False
'txt4Lac.Visible = False
'txt5Lac.Visible = False
'txt6Lac.Visible = False
'
'txtGr1.Visible = True
'txtGr2.Visible = False
'txtGr3.Visible = False
'txtGr4.Visible = False
'txtGr5.Visible = False
'txtGr6.Visible = False
'
'txtProt1.Visible = True
'txtProt2.Visible = False
'txtProt3.Visible = False
'txtProt4.Visible = False
'txtProt5.Visible = False
'txtProt6.Visible = False
'

'
'
'
'If txtCANT_DE_PARTOS.Text = (1) Then
'
'txt1Lac.Visible = True
'txt2Lac.Visible = False
'txt3Lac.Visible = False
'txt4Lac.Visible = False
'txt5Lac.Visible = False
'txt6Lac.Visible = False
'
'txtGr1.Visible = True
'txtGr2.Visible = False
'txtGr3.Visible = False
'txtGr4.Visible = False
'txtGr5.Visible = False
'txtGr6.Visible = False
'
'txtProt1.Visible = True
'txtProt2.Visible = False
'txtProt3.Visible = False
'txtProt4.Visible = False
'txtProt5.Visible = False
'txtProt6.Visible = False
'End If
'
'If txtCANT_DE_PARTOS.Text = (2) Then
'
'txt1Lac.Visible = True
'
'txt1Lac = txt1Lac.Text
'
'txt2Lac.Visible = True
'
'txt2Lac = txt2Lac.Text
'
'txt3Lac.Visible = False
'txt4Lac.Visible = False
'txt5Lac.Visible = False
'txt6Lac.Visible = False
'
'txtGr1.Visible = True
'txtGr1 = txtGr1.Text
'
'txtGr2.Visible = True
'txtGr2 = txtGr2.Text
'txtGr3.Visible = False
'txtGr4.Visible = False
'txtGr5.Visible = False
'txtGr6.Visible = False
'
'txtProt1.Visible = True
'txtProt1 = txtProt1.Text
'txtProt2.Visible = True
'txtProt2 = txtProt2.Text
'txtProt3.Visible = False
'txtProt4.Visible = False
'txtProt5.Visible = False
'txtProt6.Visible = False
'End If
'
'
'If txtCANT_DE_PARTOS.Text = (3) Then
'
'txt1Lac.Visible = True
'txt2Lac.Visible = True
'txt3Lac.Visible = True
'txt4Lac.Visible = False
'txt5Lac.Visible = False
'txt6Lac.Visible = False
'
'txtGr1.Visible = True
'txtGr2.Visible = True
'txtGr3.Visible = True
'txtGr4.Visible = False
'txtGr5.Visible = False
'txtGr6.Visible = False
'
'txtProt1.Visible = True
'txtProt2.Visible = True
'txtProt3.Visible = True
'txtProt4.Visible = False
'txtProt5.Visible = False
'txtProt6.Visible = False
'End If
'If txtCANT_DE_PARTOS.Text = (4) Then
'
'txt1Lac.Visible = True
'txt2Lac.Visible = True
'txt3Lac.Visible = True
'txt4Lac.Visible = True
'txt5Lac.Visible = False
'txt6Lac.Visible = False
'
'txtGr1.Visible = True
'txtGr2.Visible = True
'txtGr3.Visible = True
'txtGr4.Visible = True
'txtGr5.Visible = False
'txtGr6.Visible = False
'
'txtProt1.Visible = True
'txtProt2.Visible = True
'txtProt3.Visible = True
'txtProt4.Visible = True
'txtProt5.Visible = False
'txtProt6.Visible = False
'End If
'If txtCANT_DE_PARTOS.Text = (5) Then
'
'txt1Lac.Visible = True
'txt2Lac.Visible = True
'txt3Lac.Visible = True
'txt4Lac.Visible = True
'txt5Lac.Visible = True
'txt6Lac.Visible = False
'
'txtGr1.Visible = True
'txtGr2.Visible = True
'txtGr3.Visible = True
'txtGr4.Visible = True
'txtGr5.Visible = True
'txtGr6.Visible = False
'
'txtProt1.Visible = True
'txtProt2.Visible = True
'txtProt3.Visible = True
'txtProt4.Visible = True
'txtProt5.Visible = True
'txtProt6.Visible = False
'End If
'
'If txtCANT_DE_PARTOS.Text > 7 Then
'MsgBox " SE ADMITE HASTA UN MÁXIMO DE 6 PARTOS HECHOS", vbInformation
'End If
'If txtCANT_DE_PARTOS.Text = (6) Then
'
'txt1Lac.Visible = True
'txt2Lac.Visible = True
'txt3Lac.Visible = True
'txt4Lac.Visible = True
'txt5Lac.Visible = True
'txt6Lac.Visible = True
'
'txtGr1.Visible = True
'txtGr2.Visible = True
'txtGr3.Visible = True
'txtGr4.Visible = True
'txtGr5.Visible = True
'txtGr6.Visible = True
'
'txtProt1.Visible = True
'txtProt2.Visible = True
'txtProt3.Visible = True
'txtProt4.Visible = True
'txtProt5.Visible = True
'txtProt6.Visible = True
'End If
'End If
'
'
'
'If txtCANT_DE_PARTOS = (0) Then
'
'
'txt1Lac.Text = 0
' txt2Lac.Text = 0
'  txt3Lac.Text = 0
'  txt4Lac.Text = 0
'  txt5Lac.Text = 0
' txt6Lac.Text = 0
'
'txtGr1.Text = 0
' txtGr2.Text = 0
' txtGr3.Text = 0
' txtGr4.Text = 0
' txtGr5.Text = 0
'txtGr6.Text = 0
'
' txtProt1.Text = 0
' txtProt2.Text = 0
' txtProt3.Text = 0
' txtProt4.Text = 0
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'
'If txtCANT_DE_PARTOS = (1) Then
'txt1Lac.Text = txt1Lac.Text
' txt2Lac.Text = 0
'  txt3Lac.Text = 0
'  txt4Lac.Text = 0
'  txt5Lac.Text = 0
' txt6Lac.Text = 0
'
'txtGr1.Text = txtGr1.Text
' txtGr2.Text = 0
' txtGr3.Text = 0
' txtGr4.Text = 0
' txtGr5.Text = 0
'txtGr6.Text = 0
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = 0
' txtProt3.Text = 0
' txtProt4.Text = 0
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 2 Then
'txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
' txt3Lac.Text = 0
' txt4Lac.Text = 0
' txt5Lac.Text = 0
' txt6Lac.Text = 0
'
' txtGr1.Text = txtGr1.Text
' txtGr2.Text = txtGr2.Text
' txtGr3.Text = 0
' txtGr4.Text = 0
' txtGr5.Text = 0
' txtGr6.Text = 0
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = 0
' txtProt4.Text = 0
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 3 Then
'txt3Lac.Text = txt3Lac.Text
' txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
' txt4Lac.Text = 0
' txt5Lac.Text = 0
' txt6Lac.Text = 0
'
' txtGr1.Text = txtGr1.Text
' txtGr2.Text = txtGr2.Text
' txtGr3.Text = txtGr3.Text
' txtGr4.Text = 0
' txtGr5.Text = 0
' txtGr6.Text = 0
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = txtProt3.Text
' txtProt4.Text = 0
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 4 Then
' txt4Lac.Text = txt4Lac.Text
' txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
' txt4Lac.Text = txt3Lac.Text
' txt5Lac.Text = 0
' txt6Lac.Text = 0
'
'txtGr1.Text = txtGr1.Text
'txtGr2.Text = txtGr2.Text
' txtGr3.Text = txtGr3.Text
' txtGr4.Text = txtGr4.Text
' txtGr5.Text = 0
' txtGr6.Text = 0
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = txtProt3.Text
' txtProt4.Text = txtProt4.Text
' txtProt5.Text = 0
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 5 Then
'txt5Lac.Text = txt5Lac.Text
' txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
'txt4Lac.Text = txt3Lac.Text
' txt4Lac.Text = txt4Lac.Text
' txt6Lac.Text = 0
'txtGr1.Text = txtGr1.Text
' txtGr2.Text = txtGr2.Text
' txtGr3.Text = txtGr3.Text
' txtGr4.Text = txtGr4.Text
' txtGr5.Text = txtGr5.Text
' txtGr6.Text = 0
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = txtProt3.Text
' txtProt4.Text = txtProt4.Text
' txtProt5.Text = txtProt5.Text
' txtProt6.Text = 0
'End If
'If txtCANT_DE_PARTOS = 6 Then
'txt6Lac.Text = txt6Lac.Text
' txt2Lac.Text = txt2Lac.Text
' txt1Lac.Text = txt1Lac.Text
'txt3Lac.Text = txt3Lac.Text
' txt4Lac.Text = txt4Lac.Text
' txt5Lac.Text = txt5Lac.Text
'
'
' txtGr1.Text = txtGr1.Text
' txtGr2.Text = txtGr2.Text
' txtGr3.Text = txtGr3.Text
' txtGr4.Text = txtGr4.Text
' txtGr5.Text = txtGr5.Text
' txtGr6.Text = txtGr6.Text
'
' txtProt1.Text = txtProt1.Text
' txtProt2.Text = txtProt2.Text
' txtProt3.Text = txtProt3.Text
' txtProt4.Text = txtProt4.Text
' txtProt5.Text = txtProt5.Text
' txtProt6.Text = txtProt6.Text
'End If
'
'End Sub


End Sub
Private Sub txtCANT_DE_PARTOS_LostFocus()
'If txtCANT_DE_PARTOS.Text = "" Then
'Exit Sub
'End If
End Sub


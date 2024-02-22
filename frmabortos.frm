VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmabortos 
   Caption         =   "CARGA DE ABORTOS"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6840
      Top             =   2760
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmabortos.frx":0000
      OLEDBString     =   $"frmabortos.frx":0089
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CARGA_DE_ABORTOS"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MODIFICAR"
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MAS DATOS"
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BORRAR"
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmabortos.frx":0112
      Left            =   5280
      List            =   "frmabortos.frx":012B
      TabIndex        =   7
      Top             =   3240
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      Top             =   3960
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmabortos.frx":0186
      Left            =   5280
      List            =   "frmabortos.frx":0196
      TabIndex        =   5
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox txtFECHA_DEL_EVENTO 
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtRP 
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmabortos.frx":01B9
      Height          =   1815
      Left            =   6480
      TabIndex        =   2
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "frmabortos.frx":01CE
      Left            =   240
      List            =   "frmabortos.frx":01FC
      TabIndex        =   1
      Top             =   3120
      Width           =   2775
   End
   Begin MSComCtl2.MonthView MonthView1 
      Bindings        =   "frmabortos.frx":0284
      Height          =   2370
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   63700994
      CurrentDate     =   37543
   End
   Begin VB.Frame Frame1 
      Caption         =   "CONTROLES DE ABORTO"
      Height          =   1575
      Left            =   1200
      TabIndex        =   16
      Top             =   6000
      Width           =   7695
   End
   Begin VB.Label Label5 
      Caption         =   "FECHA DEL EVENTO"
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "RP"
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "RESPOSABLES"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "COMENTARIO"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "CAUSA"
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmabortos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
FRMELIMINACIONABORTOS.Show
End Sub

Private Sub Command2_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LOS DATOS DE ABORTO?", vbQuestion + vbYesNo, "CONFIRMACION")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR
With DataEnvironment1.rsCARGA_DE_ABORTOS
.Open
.AddNew
 !RP = txtRP.Text
 !FECHA_DEL_EVENTO = txtFECHA_DEL_EVENTO.Text
!CAUSA = Combo3.Text
!COMENTARIO = Combo2.Text
.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtRP = ""
Combo3 = ""
Combo2 = ""
Combo1 = ""
txtFECHA_DEL_EVENTO = ""
txtRP.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_ABORTOS
.CancelUpdate
.Close
End With

End Sub

Private Sub Command4_Click()
FRMFICHANIMAL.Show
End Sub

Private Sub Command5_Click()
MsgBox "MODIFIQUE LOS DATOS EN LA GRILLA SUPERIOR ", vbInformation, "INFORMACIÓN"
End Sub

Private Sub Form_Load()
'MsgBox (" INGRESE EL RP DEL ANIMAL POR FAVOR"), vbInformation
'MonthView1 = Date

End Sub

Private Sub List1_Click()
If List1.ListIndex = 3 Then
MsgBox "INGRESE EL RP DEL ANIMAL", vbInformation
End If

If List1.ListIndex = 0 Then
 FRMPARTOS.Show
 Else: FRMPARTOS.Hide
 If List1.ListIndex = 1 Then
 FRMSERVICIOS.Show
 Else: FRMSERVICIOS.Hide
 If List1.ListIndex = 2 Then
FRMCELO.Show
 Else: FRMCELO.Hide
 If List1.ListIndex = 4 Then
 FRMVACASECAS.Show
 Else: FRMVACASECAS.Hide
 If List1.ListIndex = 5 Then
FRMVENTAS.Show
Else: FRMVENTAS.Hide
If List1.ListIndex = 6 Then
FRMUERTES.Show
Else: FRMUERTES.Hide
If List1.ListIndex = 7 Then
frm_Medicacion_animal.Show
Else: frm_Medicacion_animal.Hide
If List1.ListIndex = 8 Then
FRMRECHAZOS.Show
Else: FRMRECHAZOS.Hide
If List1.ListIndex = 9 Then
FRMANALISIS.Show
Else: FRMANALISIS.Hide
If List1.ListIndex = 10 Then
FRMTACTORECTAL.Show
Else: FRMTACTORECTAL.Hide
If List1.ListIndex = 11 Then
FRMENFERMEDAD.Show
Else: FRMENFERMEDAD.Hide
If List1.ListIndex = 12 Then
FRMINDICACIONESESPECIALES.Show
Else: FRMINDICACIONESESPECIALES.Hide
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If





'If txtRP = "" Then
' MsgBox " Debe de ingresar el RP del animal", vbInformation, "Información"
' End If
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
'MonthView1 = Date


End Sub


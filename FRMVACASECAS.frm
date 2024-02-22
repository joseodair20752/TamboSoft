VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMVACASECAS 
   Caption         =   "INGRESO DE VACAS SECAS"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   15330
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "BORRAR"
      Height          =   495
      Left            =   8760
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   9600
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ComboBox ComboCOMENTARIO 
      Height          =   315
      ItemData        =   "FRMVACASECAS.frx":0000
      Left            =   5040
      List            =   "FRMVACASECAS.frx":0007
      TabIndex        =   11
      Top             =   6360
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "CONTROLES DE ENFERMEDAD"
      Height          =   1575
      Left            =   8640
      TabIndex        =   10
      Top             =   3840
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "FRMVACASECAS.frx":0014
      Left            =   0
      List            =   "FRMVACASECAS.frx":0042
      TabIndex        =   9
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtRP 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtFECHA_DEL_EVENTO 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox ComboRESPONSABLE 
      Height          =   315
      ItemData        =   "FRMVACASECAS.frx":00CA
      Left            =   5040
      List            =   "FRMVACASECAS.frx":00DA
      TabIndex        =   6
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MAS DATOS"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MODIFICAR"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox ComboMEDCUARTOSMAMARIOS 
      Height          =   315
      ItemData        =   "FRMVACASECAS.frx":00FD
      Left            =   5040
      List            =   "FRMVACASECAS.frx":0104
      TabIndex        =   3
      Top             =   3600
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6720
      Top             =   2520
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
      Connect         =   $"FRMVACASECAS.frx":0111
      OLEDBString     =   $"FRMVACASECAS.frx":019A
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CARGA_DE_SECADO"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FRMVACASECAS.frx":0223
      Height          =   1815
      Left            =   6240
      TabIndex        =   12
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "REGISTROS DE VACAS EN ESTADO DE SEQUEDAD"
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
   Begin MSComCtl2.MonthView MonthView1 
      Bindings        =   "FRMVACASECAS.frx":0238
      Height          =   2370
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   251592706
      CurrentDate     =   37543
   End
   Begin VB.Label Label4 
      Caption         =   "RP"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "FECHA DEL EVENTO"
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "MED. DE CUARTOS MARIOS"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "RESPONSABLE"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "COMENTARIO"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   6360
      Width           =   1695
   End
End
Attribute VB_Name = "FRMVACASECAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
FRMVACASSECASELIMINAR.Show
End Sub

Private Sub Command2_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LOS DATOS ?", vbQuestion + vbYesNo, "CONFIRMACION")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR
With DataEnvironment1.rsCARGA_DE_SECADO
.Open
.AddNew
!RP = txtRP.Text
!FECHA_DEL_EVENTO = txtFECHA_DEL_EVENTO.Text
!MEDICACION_CUARTOS_MAMARIOS = ComboMEDCUARTOSMAMARIOS.Text
!RESPONSABLE = ComboRESPONSABLE.Text
!COMENTARIO = ComboCOMENTARIO.Text
.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtRP = ""
ComboMEDCUARTOSMAMARIOS = ""
ComboRESPONSABLE = ""
ComboCOMENTARIO = ""
txtFECHA_DEL_EVENTO = ""
txtRP.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_SECADO
.CancelUpdate
.Close
End With
End Sub

Private Sub Command4_Click()
FRMFICHANIMAL.Show
End Sub

Private Sub Command5_Click()
MsgBox " MODIFIQUE LOS DATOS EN LA GRILLA SUPERIOR", vbInformation, "INFORMACIÓN"

End Sub

Private Sub Form_Load()
'MsgBox (" INGRESE EL RP DEL ANIMAL POR FAVOR"), vbInformation
'MonthView1 = Date

End Sub

Private Sub List1_Click()

If List1.ListIndex = 4 Then
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
 If List1.ListIndex = 3 Then
 frmabortos.Show
 Else: frmabortos.Hide
' If List1.ListIndex = 4 Then
' FRMVACASECAS.Show
' Else: FRMVACASECAS.Hide
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


End Sub

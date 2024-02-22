VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMTACTORECTAL 
   Caption         =   "INGRESO DE TACTOS RECTALES"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.ComboBox ComboMEDICACIONGENITAL 
      Height          =   315
      Left            =   5040
      TabIndex        =   25
      Top             =   5280
      Width           =   3495
   End
   Begin VB.ComboBox ComboENFERMEDADDEOVARIO 
      Height          =   315
      ItemData        =   "FRMTACTORECTAL.frx":0000
      Left            =   5040
      List            =   "FRMTACTORECTAL.frx":0002
      TabIndex        =   24
      Top             =   4440
      Width           =   3495
   End
   Begin VB.ComboBox ComboVIADEAPLICACION 
      Height          =   315
      ItemData        =   "FRMTACTORECTAL.frx":0004
      Left            =   5040
      List            =   "FRMTACTORECTAL.frx":0011
      TabIndex        =   21
      Top             =   5880
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   9240
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BORRAR"
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MODIFICAR"
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MAS DATOS"
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox ComboDIAGNOSTICODEUTERO 
      Height          =   315
      ItemData        =   "FRMTACTORECTAL.frx":002E
      Left            =   5040
      List            =   "FRMTACTORECTAL.frx":0047
      TabIndex        =   9
      Top             =   3000
      Width           =   3375
   End
   Begin VB.ComboBox ComboENFERMEDADDEUTERO 
      Height          =   315
      ItemData        =   "FRMTACTORECTAL.frx":00A2
      Left            =   5040
      List            =   "FRMTACTORECTAL.frx":00A4
      TabIndex        =   8
      Top             =   3720
      Width           =   3375
   End
   Begin VB.ComboBox ComboRESPONSABLE 
      Height          =   315
      ItemData        =   "FRMTACTORECTAL.frx":00A6
      Left            =   5040
      List            =   "FRMTACTORECTAL.frx":00B6
      TabIndex        =   7
      Top             =   6480
      Width           =   3495
   End
   Begin VB.TextBox txtFECHA_DEL_EVENTO 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtRP 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "FRMTACTORECTAL.frx":00D9
      Left            =   0
      List            =   "FRMTACTORECTAL.frx":0107
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
   Begin VB.ComboBox ComboCOMENTARIO 
      Height          =   315
      ItemData        =   "FRMTACTORECTAL.frx":018F
      Left            =   5040
      List            =   "FRMTACTORECTAL.frx":0196
      TabIndex        =   0
      Top             =   7080
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6120
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
      Connect         =   $"FRMTACTORECTAL.frx":01A3
      OLEDBString     =   $"FRMTACTORECTAL.frx":022C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CARGA_DE_TACTO_RECTAL"
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
      Bindings        =   "FRMTACTORECTAL.frx":02B5
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
      Caption         =   "REGISTROS DE TACTOS RECTALES"
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
      Bindings        =   "FRMTACTORECTAL.frx":02CA
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
      StartOfWeek     =   151453698
      CurrentDate     =   37543
   End
   Begin VB.Label Label9 
      Caption         =   "COMENTARIO"
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "RESPONSABLE"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "VIA DE APLICACION"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "FECHA DEL EVENTO"
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "RP"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "MEDICACION GENITAL"
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "ENFERMEDAD DE UTERO"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "DIAGNOSTICO DE UTERO "
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "ENFERMEDAD DE OVARIO"
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   4440
      Width           =   1815
   End
End
Attribute VB_Name = "FRMTACTORECTAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
FRMTACTORECTALELIMINAR.Show
End Sub

Private Sub Command2_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("�CONFIRMA EL INGRESO DE LOS DATOS DE TACTO RECTAL?", vbQuestion + vbYesNo, "CONFIRMACION")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR
With DataEnvironment1.rsCARGA_DE_TACTO_RECTAL
.Open
.AddNew
 !RP = txtRP.Text
 !FECHA_DEL_EVENTO = txtFECHA_DEL_EVENTO.Text
!DIAGNOSTICO_DE_UTERO = ComboDIAGNOSTICODEUTERO.Text
!ENFERMEDAD_DEL_UTERO = ComboENFERMEDADDEUTERO.Text
!ENFERMEDAD_DE_OVARIO = ComboENFERMEDADDEOVARIO.Text
!MEDICACION_GENITAL = ComboMEDICACIONGENITAL.Text
!VIA_DE_APLICACION = ComboVIADEAPLICACION.Text
!RESPONSABLE = ComboRESPONSABLE.Text
!COMENTARIO = ComboCOMENTARIO.Text
.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACI�N"
txtRP = ""
ComboDIAGNOSTICODEUTERO = ""
ComboENFERMEDADDEUTERO = ""
ComboENFERMEDADDEOVARIO = ""
ComboMEDICACIONGENITAL = ""
ComboVIADEAPLICACION = ""
ComboRESPONSABLE = ""
ComboCOMENTARIO = ""
txtFECHA_DEL_EVENTO = ""
txtRP.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_TACTO_RECTAL
.CancelUpdate
.Close
End With

End Sub

Private Sub Command4_Click()
FRMFICHANIMAL.Show
End Sub

Private Sub Command5_Click()
MsgBox "MODIQUE LOS DATOS EN LA GRILLA SUPERIOR", vbInformation, " INFORMACI�N"
End Sub

Private Sub Form_Load()
'MsgBox (" INGRESE EL RP DEL ANIMAL POR FAVOR"), vbInformation
'MonthView1 = Date

End Sub

Private Sub List1_Click()


If List1.ListIndex = 10 Then
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
'If List1.ListIndex = 10 Then
'FRMTACTORECTAL.Show
'Else: FRMTACTORECTAL.Hide
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

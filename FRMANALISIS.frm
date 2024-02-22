VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMANALISIS 
   Caption         =   "FORMULARIO DE ANALISIS"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   6360
      TabIndex        =   19
      Top             =   4320
      Width           =   3495
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "FRMANALISIS.frx":0000
      Left            =   6360
      List            =   "FRMANALISIS.frx":0013
      TabIndex        =   10
      Top             =   6360
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "FRMANALISIS.frx":003A
      Left            =   0
      List            =   "FRMANALISIS.frx":0068
      MultiSelect     =   1  'Simple
      TabIndex        =   9
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtRP 
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtFECHA_DEL_EVENTO 
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FRMANALISIS.frx":00F0
      Left            =   6360
      List            =   "FRMANALISIS.frx":0103
      TabIndex        =   6
      Top             =   4920
      Width           =   3495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "FRMANALISIS.frx":012D
      Left            =   6360
      List            =   "FRMANALISIS.frx":0134
      TabIndex        =   5
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BORRAR"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MAS DATOS"
      Height          =   375
      Left            =   9000
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MODIFICAR"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7080
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
      Connect         =   $"FRMANALISIS.frx":0141
      OLEDBString     =   $"FRMANALISIS.frx":01CA
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CARGA_DE_ANALISIS"
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
      Bindings        =   "FRMANALISIS.frx":0253
      Height          =   1815
      Left            =   6240
      TabIndex        =   11
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
      Caption         =   "REGISTROS DE ANALISIS"
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
      Bindings        =   "FRMANALISIS.frx":0268
      DataField       =   "FECHA_DEL_EVENTO"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   2370
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      MultiSelect     =   -1  'True
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   252248066
      CurrentDate     =   37543
   End
   Begin VB.Line Line4 
      X1              =   3960
      X2              =   9960
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      X1              =   3960
      X2              =   9960
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   3960
      Y1              =   3960
      Y2              =   7680
   End
   Begin VB.Line Line1 
      X1              =   9960
      X2              =   9960
      Y1              =   3960
      Y2              =   7680
   End
   Begin VB.Label Label4 
      Caption         =   "RP"
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "RESPONSABLE"
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "RESULTADO"
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "COMENTARIO"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "FECHA DEL EVENTO"
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "TIPO DE ANALIS"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   4200
      Width           =   1815
   End
End
Attribute VB_Name = "FRMANALISIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LOS DATOS DEL ANALISIS?", vbQuestion + vbYesNo, "CONFIRMACION")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR
With DataEnvironment1.rsCARGA_DE_ANALISIS
.Open
.AddNew
 !RP = txtRP.Text
 !FECHA_DEL_EVENTO = txtFECHA_DEL_EVENTO.Text
!TIPO_DE_ANALISIS = Combo4.Text
!COMENTARIO = Combo2.Text
!RESULTADO = Combo1.Text
!RESPONSABLE = Combo3.Text
.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtRP = ""
Combo3 = ""
Combo2 = ""
Combo1 = ""
Combo4 = ""
txtFECHA_DEL_EVENTO = ""
txtRP.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_ANALISIS
.CancelUpdate
.Close
End With
End Sub

Private Sub List1_Click()
If List1.ListIndex = 0 Then
MsgBox "INGRESE EL RP DEL ANIMAL", vbInformation
End If

'If List1.ListIndex = 0 Then
' FRMPARTOS.Show
' Else: FRMPARTOS.Hide
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

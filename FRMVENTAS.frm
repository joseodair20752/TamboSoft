VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMVENTAS 
   Caption         =   "INGRESO DE VENTAS"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   6720
      Top             =   2520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
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
      Connect         =   $"FRMVENTAS.frx":0000
      OLEDBString     =   $"FRMVENTAS.frx":0089
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CARGA_DE_VENTAS"
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
   Begin VB.CommandButton Command3 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   10440
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   11640
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BORRAR"
      Height          =   495
      Left            =   9240
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ComboBox ComboESPESIFICACION 
      Height          =   315
      ItemData        =   "FRMVENTAS.frx":0112
      Left            =   5160
      List            =   "FRMVENTAS.frx":011F
      TabIndex        =   11
      Top             =   3600
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MODIFICAR"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MAS DATOS"
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox ComboRESPONSABLE 
      Height          =   315
      ItemData        =   "FRMVENTAS.frx":0144
      Left            =   5040
      List            =   "FRMVENTAS.frx":0157
      TabIndex        =   8
      Top             =   5640
      Width           =   3495
   End
   Begin VB.TextBox txtFECHA_DEL_EVENTO 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtRP 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "FRMVENTAS.frx":0181
      Left            =   0
      List            =   "FRMVENTAS.frx":01AF
      TabIndex        =   5
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "CONTROLES DE ENFERMEDAD"
      Height          =   1215
      Left            =   8760
      TabIndex        =   4
      Top             =   4320
      Width           =   4575
   End
   Begin VB.ComboBox ComboCOMENTARIO 
      Height          =   315
      ItemData        =   "FRMVENTAS.frx":0237
      Left            =   5040
      List            =   "FRMVENTAS.frx":023E
      TabIndex        =   3
      Top             =   6360
      Width           =   3495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FRMVENTAS.frx":024B
      Height          =   1815
      Left            =   6240
      TabIndex        =   12
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      Caption         =   "REGISTROS DE VENTAS "
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "FECHA_DEL_DIA"
         Caption         =   "FECHA_DEL_DIA"
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
         DataField       =   "RP"
         Caption         =   "RP"
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
      BeginProperty Column02 
         DataField       =   "FECHA_DEL_EVENTO"
         Caption         =   "FECHA_DEL_EVENTO"
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
      BeginProperty Column03 
         DataField       =   "ESPECIFICACION"
         Caption         =   "ESPECIFICACION"
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
      BeginProperty Column04 
         DataField       =   "COMENTARIO"
         Caption         =   "COMENTARIO"
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
      BeginProperty Column05 
         DataField       =   "RESPONSABLE"
         Caption         =   "RESPONSABLE"
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
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.MonthView MonthView1 
      Bindings        =   "FRMVENTAS.frx":0260
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
      StartOfWeek     =   251854850
      CurrentDate     =   37543
   End
   Begin VB.Label Label9 
      Caption         =   "COMENTARIO"
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "RESPONSABLE"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "ESPESIFICACION"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "FECHA DEL EVENTO"
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "RP"
      Height          =   495
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FRMVENTAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
FRMVENTASELIMINAR.Show
End Sub

Private Sub Command2_Click()
Dim RESPUESTA As Integer
RESPUESTA = MsgBox("¿CONFIRMA EL INGRESO DE LA VENTA?", vbQuestion + vbYesNo, "CONFIRMACION")
If RESPUESTA = vbYes Then
On Error GoTo ERRGRABAR
With DataEnvironment1.rsCARGA_DE_VENTAS
.Open
.AddNew
!RP = txtRP.Text
!FECHA_DEL_EVENTO = txtFECHA_DEL_EVENTO.Text
!ESPECIFICACION = ComboESPESIFICACION.Text
!RESPONSABLE = ComboRESPONSABLE.Text
!COMENTARIO = ComboCOMENTARIO.Text
.Update
.Close
MsgBox " DATOS ALMACENADOS CON EXITO", vbInformation, "INFORMACIÓN"
txtRP = ""
ComboESPESIFICACION = ""
ComboRESPONSABLE = ""
ComboCOMENTARIO = ""
txtFECHA_DEL_EVENTO = ""
txtRP.SetFocus
End With
End If
Exit Sub
ERRGRABAR:
MsgBox Err.Description
With DataEnvironment1.rsCARGA_DE_VENTAS
.CancelUpdate
.Close
End With
End Sub

Private Sub Command4_Click()
FRMFICHANIMAL.Show
End Sub

Private Sub Command5_Click()
MsgBox " Modifique en la grilla presente", vbInformation, "Información"
End Sub

Private Sub Form_Load()


'MonthView1 = Date


End Sub

Private Sub List1_Click()
If List1.ListIndex = 5 Then
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

If List1.ListIndex = 6 Then
FRMUERTES.Show
Else: FRMUERTES.Hide
If List1.ListIndex = 7 Then
FRMUERTES.Show
Else: FRMUERTES.Hide
If List1.ListIndex = 8 Then
frm_Medicacion_animal.Show
Else: frm_Medicacion_animal.Hide
If List1.ListIndex = 9 Then
FRMANALISIS.Show
Else: FRMANALISIS.Hide
If List1.ListIndex = 10 Then
FRMTACTORECTAL.Show
Else: FRMTACTORECTAL.Hide
'If List1.ListIndex = 11 Then
'FRMENFERMEDAD.Show
'Else: FRMENFERMEDAD.Hide
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


End Sub

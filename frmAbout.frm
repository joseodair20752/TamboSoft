VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de Software de tambo 1.0"
   ClientHeight    =   4935
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5940
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3406.225
   ScaleMode       =   0  'User
   ScaleWidth      =   5577.967
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   0
      Top             =   3840
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Info. del sistema..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   1
      Top             =   3840
      Width           =   1485
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Este sofware est� dedicado esclusivamente para productores y estancieros solamente para el sector rural "
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SOFTWARE TAMBO 1.0"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   3735
   End
   Begin VB.Line Line5 
      X1              =   1014.176
      X2              =   4732.821
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      X1              =   4732.821
      X2              =   4732.821
      Y1              =   0
      Y2              =   414.131
   End
   Begin VB.Line Line3 
      X1              =   1014.176
      X2              =   4732.821
      Y1              =   414.131
      Y2              =   414.131
   End
   Begin VB.Line Line2 
      X1              =   1014.176
      X2              =   1014.176
      Y1              =   0
      Y2              =   414.131
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   5521.624
      Y1              =   3313.046
      Y2              =   3313.046
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   5521.624
      Y1              =   82.826
      Y2              =   82.826
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Versi�n: 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2280
      TabIndex        =   2
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Picture         =   "frmAbout.frx":00B3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9900
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opciones de seguridad de clave del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos ROOT de clave del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Const REG_DWORD = 4                      ' N�mero de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Intentar obtener ruta de acceso y nombre del programa de Info. del sistema a partir del Registro...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Intentar obtener s�lo ruta del programa de Info. del sistema a partir del Registro...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validar la existencia de versi�n conocida de 32 bits del archivo
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error: no se puede encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error: no se puede encontrar la entrada del Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "La informaci�n del sistema no est� disponible en este momento", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' C�digo de retorno
    Dim hKey As Long                                        ' Controlador de una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de datos de una clave de Registro
    Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tama�o de variable de clave de Registro
    '------------------------------------------------------------
    ' Abrir clave de registro bajo KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error de controlador...
    
    tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
    KeyValSize = 1024                                       ' Marcar tama�o de variable
    
    '------------------------------------------------------------
    ' Obtener valor de clave de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtener o crear valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agregar cadena terminada en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Encontrado valor nulo, se va a quitar de la cadena
    Else                                                    ' En WinNT las cadenas no terminan en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se ha encontrado valor nulo, s�lo se va a extraer la cadena
    End If
    '------------------------------------------------------------
    ' Determinar tipo de valor de clave para conversi�n...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Buscar tipos de datos...
    Case REG_SZ                                             ' Tipo de datos String de clave de Registro
        KeyVal = tmpVal                                     ' Copiar valor de cadena
    Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
        For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor car�cter a car�cter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a cadena
    End Select
    
    GetKeyValue = True                                      ' Se ha devuelto correctamente
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Borrar despu�s de que se produzca un error...
    KeyVal = ""                                             ' Establecer valor a cadena vac�a
    GetKeyValue = False                                     ' Fallo de retorno
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
End Function

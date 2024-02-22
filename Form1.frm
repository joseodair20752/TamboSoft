VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmBase 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Menu principal"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   10770
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3240
      Top             =   4200
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Mis documentos\SOFTWARE_DE_TAMBO1.0\BaseTambo2-0.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Mis documentos\SOFTWARE_DE_TAMBO1.0\BaseTambo2-0.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CARGA_DE_DATOS_ESTABLECIMIENTO"
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "NOMBRE_PRODUCTOR"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIRECCION"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "DIRECCION"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "LOCALIDAD"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   3240
      Width           =   6975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOCALIDAD:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "NUMERO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "NOMBRE_ESTABLECIMIENTO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NUMERO:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ESTABLECIMIENTO"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRODUCTOR"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   120
      Picture         =   "Form1.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9900
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuClave_habilitacion 
         Caption         =   "Clave de habilitacion"
      End
      Begin VB.Menu mnuCopia_de_Seguridad 
         Caption         =   "Copia de Seguridad"
      End
      Begin VB.Menu mnuApagarPc 
         Caption         =   "Apagar su PC"
      End
      Begin VB.Menu mnuSalirdelPrograma 
         Caption         =   "Salir del Programa"
      End
   End
   Begin VB.Menu MnuGeneral 
      Caption         =   "&General"
      Begin VB.Menu mnuProductor 
         Caption         =   "Productor"
      End
      Begin VB.Menu mnuTambo 
         Caption         =   "Tambo"
      End
      Begin VB.Menu mnuPersonal 
         Caption         =   "Personal"
      End
      Begin VB.Menu muTablas 
         Caption         =   "Tablas"
      End
   End
   Begin VB.Menu mnuControles 
      Caption         =   "&Controles"
      Begin VB.Menu MnuOrdeñe 
         Caption         =   "Cantidad en Tanque Primer y Segundo Ordeñe"
      End
      Begin VB.Menu mnuRegistrar_control 
         Caption         =   "Registrar control"
      End
      Begin VB.Menu mnuGestacion 
         Caption         =   "Control de Gestación"
      End
      Begin VB.Menu mnuAntiguedad 
         Caption         =   "Antiguedad de Animales "
      End
      Begin VB.Menu MnuOrdeñe1 
         Caption         =   "Grafico Primer Ordeñe"
      End
      Begin VB.Menu mnuGraf_segundo_ordeñe 
         Caption         =   "Grafico Segundo Ordeñe"
      End
      Begin VB.Menu MNUCONTROLSOMATICO 
         Caption         =   "Ingreso control Somatico"
      End
      Begin VB.Menu MNUZOMATICO 
         Caption         =   "Grafico de Evolución Somática"
      End
      Begin VB.Menu MNUCONTROL_SOMATICO 
         Caption         =   "Lista de control Somático"
      End
      Begin VB.Menu mnuGeneral_Controles 
         Caption         =   "General Controles"
      End
   End
   Begin VB.Menu MNUELIMINACIONESGENE 
      Caption         =   "&Eliminaciones Generales"
      Begin VB.Menu MNUElimina_Productor 
         Caption         =   "Eliminación de Productor"
      End
      Begin VB.Menu MNUZOMATICO2 
         Caption         =   "Eliminación ce Control Somatico"
      End
      Begin VB.Menu mnuEliminarcelos 
         Caption         =   "Eliminar celos"
      End
      Begin VB.Menu MnuEliMedi 
         Caption         =   "Eliminar Medicaciones"
      End
      Begin VB.Menu mnupartosEliminar 
         Caption         =   "Eliminación de Partos"
      End
      Begin VB.Menu mnuvacasecaseliminar 
         Caption         =   "Eliminar Vacas Secas"
      End
      Begin VB.Menu mnuServiciosEliminar 
         Caption         =   "Eliminar Servicios de Animales "
      End
      Begin VB.Menu mnurechazoseliminar 
         Caption         =   "Eliminar Vacas con Rechazos"
      End
      Begin VB.Menu MNUELIMINACIONDEVENTAS 
         Caption         =   "Eliminacion de Ventas "
      End
      Begin VB.Menu mnuEliminarvacasrectal 
         Caption         =   "Eliminar Vacas Con Tacto Rectal"
      End
      Begin VB.Menu mnuAnalisisEliminacion 
         Caption         =   "Eliminacion de Analisis de Animales"
      End
      Begin VB.Menu mnuEliminacionEnfermedad 
         Caption         =   "Eliminacion de Animales Enfermos "
      End
      Begin VB.Menu mnuVacasMuertas 
         Caption         =   "Eliminacion de Vacas Muertas"
      End
      Begin VB.Menu mnuIndicacionesespecialesEliminar 
         Caption         =   "Eliminar Animales con indicaciones especiales"
      End
      Begin VB.Menu mnuabortoseliminar 
         Caption         =   "Eliminación de Abortos"
      End
      Begin VB.Menu MNUVACASELIMINAR 
         Caption         =   "Eliminar Vacas"
      End
      Begin VB.Menu MnuSegundoOrdeñe 
         Caption         =   "Eliminación de datos Segundo Ordeñe"
      End
   End
   Begin VB.Menu mnubusqueImpreAnimal 
      Caption         =   "&Busqueda e Impresion por Animal"
      Begin VB.Menu mnuAnalisis 
         Caption         =   "BUSQUEDA POR ANALISIS"
      End
      Begin VB.Menu mnu_Animales_Indicaciones_Especiales 
         Caption         =   "BUSQUEDA DE ANIMALES CON INDICACIONES ESPECIALES"
      End
      Begin VB.Menu Mnu_Medicacion 
         Caption         =   "BUSQUEDA DE ANIMALES CON MEDICACION"
      End
      Begin VB.Menu MNUPRIMER_ORDEÑE 
         Caption         =   "BUSQUEDA DE DATOS DEL PRIMER ORDEÑE"
      End
      Begin VB.Menu MNUSEGUNDO_ORDEÑE 
         Caption         =   "BUSQUEDA DE DATOS DEL SEGUNDO ORDEÑE"
      End
      Begin VB.Menu MnuCelos 
         Caption         =   "BUSQUEDA DE ANIMALES CON CELOS "
      End
      Begin VB.Menu MNU_ZOMATICO 
         Caption         =   "BUSQUEDA DE PARAMETROS SOMATICOS"
      End
      Begin VB.Menu MNUBUS_ABORTOS 
         Caption         =   "BUSQUEDA DE ANIMALES CON ABORTOS"
      End
      Begin VB.Menu FORM1 
         Caption         =   "BUSQUEDA DE ANIMALES POR EDAD"
         Shortcut        =   ^E
      End
      Begin VB.Menu MNUANIMALES_ENFERMOS 
         Caption         =   "BUSQUEDA DE ANIMALES ENFERMOS "
      End
      Begin VB.Menu mnuBusvacas 
         Caption         =   "BUSQUEDA DE VACAS"
      End
   End
   Begin VB.Menu MNUANIMALES 
      Caption         =   "&Animales"
      Index           =   1
      Begin VB.Menu mnuVacas_Ingreso 
         Caption         =   "Ingreso de vacas"
      End
      Begin VB.Menu mnuCarga_de_Eventos 
         Caption         =   "Carga de Eventos"
      End
      Begin VB.Menu mnuvaquillonas 
         Caption         =   "Ingreso de Vaquillonas"
      End
      Begin VB.Menu mnuFicha 
         Caption         =   "Ficha"
      End
      Begin VB.Menu mnuLoteo 
         Caption         =   "Loteo"
      End
      Begin VB.Menu mnuToros 
         Caption         =   "Toros"
         Begin VB.Menu mnuStock_Semen 
            Caption         =   "Stock de Semen"
         End
         Begin VB.Menu mnuConsultar 
            Caption         =   "Consultar Padrón"
         End
         Begin VB.Menu mnuActualizar_Padrón 
            Caption         =   "Actualizar Padrón"
         End
      End
   End
   Begin VB.Menu mnuListados 
      Caption         =   "&Listados"
      Begin VB.Menu mnuGeneral_Animales 
         Caption         =   " Listados Generales de Animales"
         Begin VB.Menu mnuProduccion 
            Caption         =   "Produccion"
            Index           =   1
         End
         Begin VB.Menu mnuTorosListado 
            Caption         =   "Listado de Toros"
         End
         Begin VB.Menu MnuGeneral_servicios 
            Caption         =   "Listado General de servicios"
         End
         Begin VB.Menu mnulistgeneral 
            Caption         =   "Listados de Vacas en general"
         End
         Begin VB.Menu mnuMedi 
            Caption         =   "Listado de animales medicados "
         End
         Begin VB.Menu MnuEnfermos 
            Caption         =   "Listado de animales enfermos "
         End
         Begin VB.Menu MnuCelos_Lista 
            Caption         =   "Listado de Celos"
         End
         Begin VB.Menu MnuList_Especiales 
            Caption         =   "Listado de animales con Indicaciones especiales"
         End
         Begin VB.Menu mnuAnálisis 
            Caption         =   "Listado de análisis"
         End
         Begin VB.Menu mnuTerneros 
            Caption         =   "Listado de Terneros "
         End
         Begin VB.Menu mnuAbortos 
            Caption         =   "Listado de abortos"
         End
      End
      Begin VB.Menu mnuIndicativos_animales 
         Caption         =   "Indicativos de Animales"
         Begin VB.Menu mnuVacasSecasPreñez 
            Caption         =   "Lista de Vacas Secas"
         End
         Begin VB.Menu mnuVacasParaTacto 
            Caption         =   "Lista de tactos Rectales"
         End
         Begin VB.Menu mnuVacasaparir 
            Caption         =   "Lista de Partos"
         End
         Begin VB.Menu mnuvaquillonas1 
            Caption         =   "Listado de vaquillonas"
         End
         Begin VB.Menu mnuVacasconIndicacionderechazo 
            Caption         =   "Vacas con Indicacion de rechazo"
         End
      End
      Begin VB.Menu mnuOperativos 
         Caption         =   "Operativos"
         Begin VB.Menu mnuProducción 
            Caption         =   "Cantidad de Vacas en Producción"
         End
      End
   End
   Begin VB.Menu mnuAcercade 
      Caption         =   "&Acerca de..."
      Begin VB.Menu mnuIralapaginadeConaprole 
         Caption         =   "Ir a la pagina de Conaprole"
      End
      Begin VB.Menu mnuProducciondelPrograma 
         Caption         =   "Produccion del programa"
      End
   End
End
Attribute VB_Name = "frmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub FORM1_Click()
Form1_EDAD_ANIMAL.Show
End Sub

Private Sub mnu_Animales_Indicaciones_Especiales_Click()
frm_INDICACIONES_ESPECIALES.Show
End Sub

Private Sub Mnu_Medicacion_Click()
FRM_MEDICACION.Show
End Sub

Private Sub MNU_ZOMATICO_Click()

FRMCONTROLSOMATICO.Show
End Sub



Private Sub mnuAbortos_Click()
RptListadoGeneralAbortos.Show
End Sub

Private Sub mnuabortoseliminar_Click()
FRMELIMINACIONABORTOS.Show
End Sub

Private Sub mnuAnalisis_Click()
Form2.Show
End Sub

Private Sub mnuAnálisis_Click()
FrmListadoAnalisis.Show

End Sub

Private Sub mnuAnalisisEliminacion_Click()
FRMELIMINACIONANALISIS.Show
End Sub

Private Sub MNUANIMALES_ENFERMOS_Click()
FRM_PARAMETRO_ENFERMEDADES.Show
End Sub

Private Sub mnuAntiguedad_Click()
frmedad_Animal.Show
'MsgBox " Esto es un demo de SOFTWARE DE TAMBO 1.0", vbInformation
End Sub

Private Sub mnuApagarPc_Click()
'FORM1.Show
End Sub

Private Sub MNUBUS_ABORTOS_Click()
FRM_ABORTOS.Show
End Sub

Private Sub mnuBusvacas_Click()
FRMPAR1.Show
End Sub

Private Sub mnuCarga_de_Eventos_Click()
FRMCELO.Show
End Sub

Private Sub MNUCELOS_Click()
FRM_CELOS_PARAMETRO.Show
End Sub

Private Sub mnuCelulasSomáticasUltimosControles_Click()

End Sub

Private Sub MnuCelos_Lista_Click()
FrmListado_Celos.Show
End Sub

Private Sub mnuClave_habilitacion_Click()
frmHabilitacion.Show
End Sub

Private Sub mnuConfiguracion_Click()

End Sub

Private Sub MNUCONTROL_SOMATICO_Click()
RPTLISTA_GENERAL_DE_CONTROL_ZOMATICO.Show
End Sub

Private Sub MNUCONTROLSOMATICO_Click()
FRMINGRESOCONTROLSOMATICO.Show
End Sub

Private Sub mnuCopia_de_Seguridad_Click()
FRMBACKUP.Show
End Sub

Private Sub MnuEliMedi_Click()
FRMELIMINACIONMEDICACION.Show
End Sub

Private Sub MNUElimina_Productor_Click()
FRMELIMINACIONPRODUCTOR.Show
End Sub

Private Sub MNUELIMINACIONDEVENTAS_Click()
FRMVENTASELIMINAR.Show
End Sub

Private Sub mnuEliminacionEnfermedad_Click()
FRMELIMINACIONENFERMEDAD.Show
End Sub

Private Sub mnuEliminarcelos_Click()
FRMELIMIANCIONCELOS.Show
End Sub

Private Sub mnuEliminarvacasrectal_Click()
FRMTACTORECTALELIMINAR.Show
End Sub

Private Sub mnuEntidad_Click()

End Sub

Private Sub mnuEventos_Click()

End Sub

Private Sub MnuEnfermos_Click()
FRMLISTADO_ENFERMOS.Show
End Sub

Private Sub mnuFicha_Click()
FRMFICHANIMAL.Show
End Sub

Private Sub mnuGeneral_Controles_Click()
Dim RESPUESTA As Integer
Dim RESPUESTA2 As Integer
RESPUESTA = MsgBox("¿ DESEA INGRESAR AL PRIMER ORDEÑE?", vbYesNo)
RESPUESTA2 = MsgBox("¿DESEA INGRESAR AL SEGUNDO ORDEÑE?", vbYesNo)


If RESPUESTA = vbYes Then
FRMORDEÑE1.Show
If RESPUESTA2 = vbYes Then

FRMSEGUNDORDEÑE.Show
End If
End If
End Sub

Private Sub mnuGenerla_Animales_Click()
frmIngresosdeVacas.Show
End Sub

Private Sub MnuGeneral_servicios_Click()
Rpt_Listado_General_de_Servicios.Show
End Sub

Private Sub mnuGestacion_Click()
FRMGESTACION.Show
End Sub

Private Sub mnuGraf_segundo_ordeñe_Click()
FRMCONT_SEGUNDO_ORDEÑE.Show
End Sub

Private Sub mnuIndicacionesespecialesEliminar_Click()
FRMINDICACIONESPECIALELIMINAR.Show
End Sub

Private Sub mnuParametrosReproductivos_Click()

End Sub

Private Sub mnuIralapaginadeConaprole_Click()
frm_Conaprole.Show
End Sub

Private Sub MnuList_Especiales_Click()
FrmListEspeciales.Show
End Sub

Private Sub mnulistgeneral_Click()
frmListVacasgeneral.Show
End Sub

Private Sub mnuLoteo_Click()
FRMLOTEO.Show
End Sub

Private Sub mnuMedi_Click()
FrmMedicados.Show
End Sub

Private Sub MnuOrdeñe_Click()
frmlistado_en_taque_de_Leche.Show
End Sub

Private Sub MNUORDEÑE1_Click()
FRMTANQUE_GRAFICO.Show
End Sub

Private Sub mnupartosEliminar_Click()
FRMELIMIANACIONPARTOS.Show
End Sub

Private Sub MNUPRIMER_ORDEÑE_Click()
FRM_PRIMER_ORDEÑE.Show
End Sub

Private Sub mnuProduccion_Click(Index As Integer)
RptListado_General_de_Vacas_en_Produccion.Show
End Sub

Private Sub mnuProducción_Click()
Form3.Show
End Sub

Private Sub mnuProducciondelPrograma_Click()
AGRO_SOFT.Show
End Sub

Private Sub mnuProductor_Click()
FRMDATOSESTABLECIMIENTO.Show
End Sub

Private Sub mnurechazoseliminar_Click()
FRMRECHAZOSELIMINAR.Show
End Sub

Private Sub mnuRegistrar_control_Click()
FRMINSPECTORES.Show
End Sub

Private Sub mnuResumenesdeControles_Click()

End Sub

Private Sub mnuSalirdelPrograma_Click()
End
End Sub

Private Sub MNUSEGUNDO_ORDEÑE_Click()
FRM_SEGUNDO_ORDEÑE.Show
End Sub

Private Sub MnuSegundoOrdeñe_Click()
FrmEliminación_Segundo_Ordeñe.Show
End Sub

Private Sub mnuServiciosEliminar_Click()
FRMSERVICIOSELIMINAR.Show
End Sub

Private Sub mnuStock_Semen_Click()
FRMTOROS.Show
End Sub

Private Sub PaySlide1_GotFocus()

End Sub

Private Sub mnuTambo_Click()
frm_Tambo.Show
End Sub

Private Sub mnuVacas_Click(Index As Integer)

End Sub

Private Sub mnuTerneros_Click()
RptListado_de_Terneros.Show
End Sub

Private Sub mnuTorosListado_Click()
RptListado_General_de_Toros.Show
End Sub

Private Sub mnuVacas_Ingreso_Click()
frmIngresosdeVacas.Show
'MsgBox " Esto es un demo de SOFTWARE DE TAMBO 1.0", vbInformation
End Sub

Private Sub mnuVacasaparir_Click()
RptListado_General_de_Partos.Show
End Sub

Private Sub mnuVacasconIndicacionderechazo_Click()
Rpt_Listado_General_Rechazos.Show
End Sub

Private Sub mnuvacasecaseliminar_Click()
FRMVACASSECASELIMINAR.Show
End Sub

Private Sub MNUVACASELIMINAR_Click()
FRMELIMINACIÓNVACAS.Show
End Sub

Private Sub mnuVacasMuertas_Click()
FRMELIMINACIONVACASMUERTAS.Show
End Sub

Private Sub mnuVacasParaTacto_Click()
RptListado_general_de_TactosRectales.Show
End Sub

Private Sub mnuVacasSecasPreñez_Click()
RptListado_General_Vacas_Secas.Show
End Sub

Private Sub MnuVaquillonas_Click()
FRMVAQUILLONAS.Show
End Sub

Private Sub mnuvaquillonas1_Click()
frmListVaquillonas.Show
End Sub

Private Sub MNUZOMATICO_Click()
FRMZOMATICO.Show
End Sub

Private Sub MNUZOMATICO2_Click()
FRMCONTROLSOMATICO.Show
End Sub

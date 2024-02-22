VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim CAT As Catalog
Dim CMD As Command

Set CAT = New Catalog
Set CMD = New Command
CAT.ActiveConnection = "PROVIDER= MICROSOFT.JET.OLEDB.4.0;Data Source=C:\Mis documentos\Software de tambo\BaseTambo.mdb"
CMD.CommandText = "PARAMETERS RP Value; SELECT [CARGA_DE_VACAS].[RP], [CARGA_DE_VACAS].[NOMBRE_MADRE], [CARGA_DE_VACAS].[NOMBRE_PADRE], [CARGA_DE_VACAS].[PROCEDENCIA], [CARGA_DE_VACAS].[AÑOS], [CARGA_DE_VACAS].[MESES] From CARGA_DE_VACAS WHERE ((([CARGA_DE_VACAS].[RP])=[RP]))"
CAT.Procedures.Append "UNEMPLEADO", CMD

Set CAT = Nothing
Set CMD = Nothing





End Sub

Attribute VB_Name = "Module1"
Global base As New ADODB.Connection
Sub main()
With base
    .CursorLocation = adUseClient
    .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BaseTambo2-0.mdb;Persist Security Info=False"
    MsgBox "Usted esta conectado a la base de datos del programa", vbInformation, "Conectado"
    frmBase.Show
   End With

End Sub

Attribute VB_Name = "Module1"
Public CON As New ADODB.Connection
Public RS As New ADODB.Recordset
Public Sub frm()
If CON.State = 1 Then CON.Close
 CON.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & App.Path & "/BANK1.MDB"
End Sub

Attribute VB_Name = "Module1"
Public con As ADODB.Connection
Public rs As ADODB.Recordset
Public userLog As String

Public Sub connect()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.CursorLocation = adUseClient
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\data.mdb"
con.Open
End Sub

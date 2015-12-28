<%
Class Connection
	Public con
	
	Private Sub Close
    con.Close
    Set con = Nothing
  End Sub
	
	Private Sub Class_Initialize
		Set con = Server.CreateObject("adodb.connection")
		con.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(".") & "\db\db1.mdb;"
  End Sub

  Private Sub Class_Terminate
    Close
  End Sub
	
	Public Function Execute(byval sql)
		Set Execute = con.Execute(sql)
	End Function
End Class
%>
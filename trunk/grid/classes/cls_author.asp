<%
Class Author
	Public Id
	Public LastName
	Public FirstName
	
	Public Sub Init(r)
		Id = r("Id")
		LastName = r("LastName")
		FirstName = r("FirstName")
	End Sub
	
	Public Sub Lookup
		Dim con
		Dim rs
		If Id <> "" Then
			Set con = New Connection
			Set rs = con.Execute("select * from authors where id=" & Id)
			If Not rs.Eof Then Init(rs)
			rs.Close
			Set rs = Nothing
			Set con = Nothing
		End If
	End Sub
	
	Public Function getAuthorList
		Set getAuthorList = New List
		getAuthorList.Init "Author", "select * from authors order by id"
	End Function
	
	Public Function Add
		Dim con
		Add = ""
		On Error Resume Next
			Set con = New Connection
			con.Execute("insert into authors (lastname, firstname) values (" & quote(LastName) & "," & quote(FirstName) & ")")
			Id = con.Execute("select @@identity")(0)
			ec = Err.Number
			ed = Err.Description
		On Error Goto 0
		If ec = 0 Then
			Add = "OK"
		Else
			Add = "Author. Error: " & ed
		End If
	End Function
	
	Public Property Get FullName
		FullName = LastName & ", " & FirstName
	End Property
	
End Class
%>
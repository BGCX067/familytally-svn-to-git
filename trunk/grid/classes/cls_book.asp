<%
Function IfElse (Val, Iftrue, Ifnot)
	IfElse = Ifnot
	If CBool(Val) Then 
		IfElse = Iftrue
	End If
End Function

Class Book
	Public Id
	Public Instore
	Public Name
	Public AuthorId
	Public PubDate
	Public Price
	
	Public Sub Init(r)
		Id = r("Id")
		On Error Resume Next
			Dim values
			values = Split(r("values"),"|")
			If isEmpty(r("Instore")) Then Instore = IfElse(values(0),1,0) Else Instore = IfElse(r("Instore"),1,0)
			If isEmpty(r("Name")) Then Name = values(1) Else Name = r("Name")
			If isEmpty(r("AuthorId")) Then AuthorId = values(2) Else AuthorId = r("AuthorId")
			If isEmpty(r("PubDate")) Then PubDate = values(3) Else PubDate = r("PubDate")
			If isEmpty(r("Price")) Then Price = values(4) Else Price = r("Price")
		
		On Error Goto 0
		
	End Sub
	
	Public Sub Lookup
		Dim con
		Dim rs
		If Id <> "" Then
			Set con = New Connection
			Set rs = con.Execute("select * from books where id=" & Id)
			If Not rs.Eof Then Init(rs)
			rs.Close
			Set rs = Nothing
			Set con = Nothing
		End If
	End Sub
	
	Public Property Get AuthorName
		Dim con
		Dim rs
		AuthorName = ""
		If AuthorId <> "" Then
			Set a = New Author
			a.Id = AuthorId
			a.Lookup
			AuthorName = a.LastName & ", " & a.FirstName
			Set a = Nothing
		End If
	End Property
	
	Public Function getBookList
		Set getBookList = New List
		getBookList.Init "Book", "select * from books order by name"
	End Function
	
	Public Function Add
		Dim con
		Add = ""
		On Error Resume Next
			Set con = New Connection
			con.Execute("insert into books (instore,name, authorid, pubdate, price) values("&Instore&","&quote(Name)&","&AuthorId&","&quote(PubDate)&","&Price&")")
			 '(null, null, null, 0)")
			Id = con.Execute("select @@identity")(0)
			ec = Err.Number
			ed = Err.Description
		On Error Goto 0
		If ec = 0 Then
			Add = "OK"
		Else
			Add = "Book. Error: " & ed
		End If
	End Function
	
	Public Function Update
		Dim con
		Update = ""
		On Error Resume Next
			If isNull(AuthorId) OR isEmpty(AuthorId) OR AuthorId = "" Then AuthorId = "null"
			If jsnquote(Price) > 0 Then Price = Replace(FormatNumber(Price, 2), ",", ".")
			Set con = New Connection
			'Response.Write("update books set instore=" & Instore & ", name=" & quote(Name) & ", authorid=" & AuthorId & ", pubdate=" & quote(PubDate) & ", price=" & Price & " where id=" & Id)
			con.Execute("update books set instore=" & Instore & ", name=" & quote(Name) & ", authorid=" & AuthorId & ", pubdate=" & quote(PubDate) & ", price=" & Price & " where id=" & Id)
			ec = Err.Number
			ed = Err.Description
		On Error Goto 0
		If ec = 0 Then
			Update = "OK"
		Else
			Update = "Book. Error: " & ed & "update books set instore=" & quote(Instore) & ", name=" & quote(Name) & ", authorid=" & AuthorId & ", pubdate=" & quote(PubDate) & ", price=" & Price & " where id=" & Id
		End If
	End Function
	
	Public Function Delete
		Dim con
		Delete = ""
		On Error Resume Next
			Set con = New Connection
			con.Execute("delete from books where id=" & Id)
			ec = Err.Number
			ed = Err.Description
		On Error Goto 0
		If ec = 0 Then
			Delete = "OK"
		Else
			Delete = "Book. Error: " & ed
		End If
	End Function
	
End Class
%>
<!-- #include file="include/header.asp" -->
<%
	Dim a, b
	Dim action_type, arr, res, bNewAuthor
	
	Set b = New Book
	b.Id = jsnquote(Request("id"))
	
	action_type = Request("type")
	If action_type = "add" Then
		b.Init(Request)
		res = b.Add
	ElseIf action_type = "update" Then
		b.Lookup
		b.Init(Request)
		
		bNewAuthor = False
		' add new author
		If jsnquote(b.AuthorId) = 0 AND b.AuthorId <> "" Then
			arr = Split(b.AuthorId, ",")
			If UBound(arr) > -1 Then
				Set a = New Author
				a.LastName = arr(0)
				If UBound(arr) > 0 Then a.FirstName = arr(1) Else a.FirstName = Null
				res = a.Add
				b.AuthorId = a.Id
				bNewAuthor = True
			Else
				res = "Author. Error: Incorrect author name"
			End If
		Else
			res = "OK"
		End If
		
		If res = "OK" Then res = b.Update
	ElseIf action_type = "delete" Then
		res = b.Delete
	End If
%>
<%
Response.ContentType = "text/xml"
%>
<status value="ok" oldid="<%=Request("id")%>" rowid="<%=b.Id%>"/>

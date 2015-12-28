<!-- #include file="include/header.asp" -->
<%
	Dim a, b
	Dim aList, bList
	
	Response.ContentType = "text/xml"
	Response.Write "<?xml version=""1.0"" encoding=""iso-8859-1""?>" & vbCrLf
	Response.Write "<rows>" & vbCrLf
	
	'Response.Write "<head>" & vbCrLf
	'Response.Write "<column width=""50"" type=""ch"" align=""center"" color=""#f7f6f0"" sort=""str"">In store</column>" & vbCrLf
	'Response.Write "<column width=""250"" type=""ed"" align=""left"" color=""#ffffff"" sort=""str"">Book Title</column>" & vbCrLf
	'Response.Write "<column width=""170"" type=""co"" align=""left"" color=""#f7f6f0"" sort=""str"">Author" & vbCrLf
	
	'Set aList = (New Author).getAuthorList
	'Do
	'	Set a = aList.NextItem
	'	If a is Nothing Then Exit Do
	'	Response.Write "<option value=""" & a.Id & """>" & a.FullName & "</option>" & vbCrLf
	'Loop
	'Set aList = Nothing
	
	'Response.Write "</column>" & vbCrLf
	'Response.Write "<column width=""70"" type=""ed"" align=""center"" color=""#ffffff"" sort=""int"">Year</column>" & vbCrLf
	'Response.Write "<column width=""50"" type=""price"" align=""right"" color=""#ffffff"" sort=""str"">Price</column>" & vbCrLf
	'Response.Write "</head>" & vbCrLf
	
	Set bList = (New Book).getBookList
	Do
		Set b = bList.NextItem
		If b is Nothing Then Exit Do
		Response.Write "<row id=""" & b.Id & """>" & vbCrLf
		Response.Write "<cell>"& b.Instore &"</cell>" & vbCrLf
		Response.Write "<cell>" & b.Name & "</cell>" & vbCrLf
		Response.Write "<cell>" & b.AuthorId & "</cell>" & vbCrLf
		Response.Write "<cell>" & b.PubDate & "</cell>" & vbCrLf
		Response.Write "<cell>" & Replace(FormatNumber(b.Price, 2), ",", ".") & "</cell>" & vbCrLf
		Response.Write "</row>" & vbCrLf
	Loop
	Set bList = Nothing
	
	Response.Write "</rows>" & vbCrLf
%>

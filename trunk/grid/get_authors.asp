<!-- #include file="include/header.asp" -->
<%
	Response.ContentType = "text/xml"
	Response.Write "<?xml version=""1.0"" encoding=""iso-8859-1""?>" & vbCrLf
	Response.Write "<authors>" & vbCrLf
	Set aList = (New Author).getAuthorList
	Do
		Set a = aList.NextItem
		If a is Nothing Then Exit Do
		Response.Write "<author value=""" & a.Id & """>" & a.FullName & "</author>" '& vbCrLf
	Loop
	Set aList = Nothing
	Response.Write "</authors>" & vbCrLf
%>
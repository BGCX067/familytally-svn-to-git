<%
	Function jsnquote(byval s)
		If isNull(s) OR isEmpty(s) OR s = "" Then s = 0 Else s = NCDbl(s)
		jsnquote = s
	End Function
	
	Function NCDbl(byval s)
	  If IsNull(s) Then
	    NCDbl = Null
	  Else
  		On Error Resume Next
  	    NCDbl = CDbl (s)
  	  On Error Goto 0
	  End If 
	End Function
	
	Function quote(byval s)
		If isEmpty(s) OR isNull(s) Then s = "NULL" Else s = "'" & replace(s,"'","''") & "'"
		quote = s
	End Function
%>
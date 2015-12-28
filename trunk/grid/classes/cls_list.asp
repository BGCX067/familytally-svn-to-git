<%
Class List
  Private con
  Private rs
  Private className
	
  Public Sub Init (in_className, sql)
    className = in_className
    Set con = New Connection
  	Set rs = con.Execute(sql)
  End Sub
	
	Public Function NextItem
		Set NextItem = Nothing
		
		if (Not isEmpty(rs)) then
			If Not rs.EOF Then
				on error resume next
				Execute "Set NextItem = New " & className
				Response.Write("0000")
				NextItem.Init(rs)
				rs.MoveNext
				ec = Err.Number
				ed = Err.Description
				on error goto 0
				if (ec <> 0) Then Wr "List error (see cls_list.asp): " & ec & ", " & ed
			End If
		End If
	End Function
	
	Public Sub MoveFirst
		on error resume next
		rs.MoveFirst
		on error goto 0
	End Sub
	
  Private Sub Close
  	if (not isEmpty(rs)) then
	    rs.Close
		end if
		
    Set rs = Nothing
    Set con = Nothing
  End Sub
	
  Private Sub Class_Terminate
    Close
  End Sub
	
End Class
%>
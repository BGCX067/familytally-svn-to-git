<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ͥ�¶�֧�������</title>
<!-- calendar stylesheet -->
<link rel="stylesheet" type="text/css" media="all" href="calendar/calendar-win2k-cold-1.css" title="win2k-cold-1" />

<!-- main calendar program -->
<script type="text/javascript" src="calendar/calendar.js"></script>

<!-- language for the calendar -->
<script type="text/javascript" src="calendar/calendar_zh.js"></script>

<!-- the following script defines the Calendar.setup helper function, which makes
   adding a calendar a matter of 1 or 2 lines of code. -->
<script type="text/javascript" src="calendar/calendar-setup.js"></script>
<script language="JavaScript">
<!-- start hiding
function check_in(obj)
{
	var jh_zc = obj.JH_ZC.value;
	var sj_zc = obj.SJ_ZC.value;
	flag1 = isNaN(jh_zc);
	flag2 = isNaN(sj_zc);
	if(flag1||flag2)
	{
		alert("����������!");
		if(flag2){obj.SJ_ZC.select();obj.SJ_ZC.focus();}
		if(flag1){obj.JH_ZC.select();obj.JH_ZC.focus();}
		return false;
	}
	else
	{
		//alert("����!");
		//setfocus(obj.list_num);
		return true;
	}
}
// stop hiding -->
</script>
</head>
<body>
<%
Set conn = Server.CreateObject ("ADODB.Connection")
conn.open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=e:\web\��ͥ��֧��.accdb;Persist Security Info=False"
function urldecoding(vstrin) 
	dim i,strreturn,strSpecial 
	strSpecial = "!""#$%&'()*+,./:;<=>?@[\]^`{|}~%" 
	strreturn = "" 
	for i = 1 to len(vstrin) 
		thischr = mid(vstrin,i,1) 
		if thischr="%" then 
			intasc=eval("&h"+mid(vstrin,i+1,2)) 
			if instr(strSpecial,chr(intasc))>0 then 
				strreturn= strreturn & chr(intasc) 
				i=i+2 
			else 
				intasc=eval("&h"+mid(vstrin,i+1,2)+mid(vstrin,i+4,2)) 
				strreturn= strreturn & chr(intasc) 
				i=i+5 
			end if 
		else 
			if thischr="+" then 
				strreturn= strreturn & " " 
			else 
				strreturn= strreturn & thischr 
			end if 
		end if 
	next 
	urldecoding = strreturn 
end function 

if request.form("zc_add")<>"" and cbool(request.form("zc_add")) then
	jh_zc=trim(replace(request.form("jh_zc"),"'","''"))
	sj_zc=trim(replace(request.form("sj_zc"),"'","''"))
	bz=trim(replace(request.form("bz"),"'","''"))
	zc_rq=trim(replace(request.form("zc_rq"),"'","''"))
	zcr=trim(replace(request.form("zcr"),"'","''"))
	syr=trim(replace(request.form("syr"),"'","''"))
	txr=trim(replace(request.form("txr"),"'","''"))
	if request.form("zc_xm")="" or isnull(request.form("zc_xm")) then
		response.Write("����ѡ��֧����Ŀ�˰�")
		response.End()
	else
		zc_xm=request.form("zc_xm")
	end if
	sql_add="insert into ��ͥ��֧���굥(jh_zc,sj_zc,bz,zc_rq,zc_xm,zcr,syr,txr) values('"&jh_zc&"','"&sj_zc&"','"&bz&"','"&zc_rq&"',"&zc_xm&",'"&zcr&"','"&syr&"','"&txr&"')"
	'response.Write(sql_add)
	'response.End()
	set cmd=Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection=conn
	cmd.CommandText = sql_add
	cmd.Execute


end if
%>
<table align="center" cellpadding="2" cellspacing="1" border="0" bgcolor="#5780B4" width="100%">
  <tr>
  <%
	set rs_by=server.CreateObject("adodb.recordset")
	sql = "SELECT * From ���¼ƻ�����" 
	rs_by.open sql,conn
	If rs_by("by_jh_sr")="" Or IsNull(rs_by("by_jh_sr")) Then 
		by_jh_sr=0
	Else
		by_jh_sr=rs_by("by_jh_sr")
	End if
  %>
    <td bgcolor="#B8CCE4">���¼ƻ�����</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(by_jh_sr,,-1)%></td>
  <%
  	rs_by.close
	sql = "SELECT * From ���¼ƻ�֧��" 
	rs_by.open sql,conn
	If rs_by("by_jh_zc")="" Or IsNull(rs_by("by_jh_zc")) Then 
		by_jh_zc=0
	Else
		by_jh_zc=rs_by("by_jh_zc")
	End if
  %>
    <td bgcolor="#B8CCE4">���¼ƻ�֧��</td>
	<td bgcolor="#DBE5F1"><%=FormatCurrency(by_jh_zc,,-1)%></td>
  <%
  	rs_by.close
	sql = "SELECT * From ���¼ƻ���ծ" 
	rs_by.open sql,conn
	If rs_by("by_jh_fz")="" Or IsNull(rs_by("by_jh_fz")) Then 
		by_jh_fz=0
	Else
		by_jh_fz=rs_by("by_jh_fz")
	End if
  %>
    <td bgcolor="#B8CCE4">���¼ƻ���ծ</td>
	<td bgcolor="#DBE5F1"><%=FormatCurrency(by_jh_fz,,-1)%></td>
  </tr>
  <tr>
  <%
  	rs_by.close
	sql = "SELECT * From ����ʵ������" 
	rs_by.open sql,conn
	If rs_by("by_sj_sr")="" Or IsNull(rs_by("by_sj_sr")) Then 
		by_sj_sr=0
	Else
		by_sj_sr=rs_by("by_sj_sr")
	End if
  %>
    <td bgcolor="#B8CCE4">����ʵ������</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(by_sj_sr,,-1)%></td>
  <%
  	rs_by.close
	sql = "SELECT * From ����ʵ��֧��" 
	rs_by.open sql,conn
	If rs_by("by_sj_zc")="" Or IsNull(rs_by("by_sj_zc")) Then 
		by_sj_zc=0
	Else
		by_sj_zc=rs_by("by_sj_zc")
	End if
  %>
    <td bgcolor="#B8CCE4">����ʵ��֧��</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(by_sj_zc,,-1)%></td>
  <%
  	rs_by.close
	sql = "SELECT * From ����ʵ�ʸ�ծ" 
	rs_by.open sql,conn
	If rs_by("by_sj_fz")="" Or IsNull(rs_by("by_sj_fz")) Then 
		by_sj_fz=0
	Else
		by_sj_fz=rs_by("by_sj_fz")
	End if
  %>
    <td bgcolor="#B8CCE4">����ʵ�ʸ�ծ</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(by_sj_fz,,-1)%></td>
  </tr>
  <tr>
    <td bgcolor="#B8CCE4">����������</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(by_sj_sr-by_jh_sr,,-1)%></td>
    <td bgcolor="#B8CCE4">����֧�����</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(by_sj_zc-by_jh_zc,,-1)%></td>
    <td bgcolor="#B8CCE4">���¸�ծ���</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(by_sj_fz-by_jh_fz,,-1)%></td>
  </tr>
  <%
  	rs_by.close
	set rs_by=nothing
  %>
  <tr>
    <td height="0" bgcolor="#5780B4"></td>
    <td height="0" bgcolor="#5780B4"></td>
    <td height="0" bgcolor="#5780B4"></td>
    <td height="0" bgcolor="#5780B4"></td>
    <td height="0" bgcolor="#5780B4"></td>
    <td height="0" bgcolor="#5780B4"></td>
  </tr>
  <tr>
  <%
	set rs_sy=server.CreateObject("adodb.recordset")
	sql = "SELECT * From ���¼ƻ�֧��" 
	rs_sy.open sql,conn
	sy_jh_zc=rs_sy("sy_jh_zc")
  %>
    <td bgcolor="#B8CCE4">���¼ƻ�֧��</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(sy_jh_zc,,-1)%></td>
  <%
  	rs_sy.close
	sql = "SELECT * From ����ʵ��֧��" 
	rs_sy.open sql,conn
	sy_sj_zc=rs_sy("sy_sj_zc")
  %>
    <td bgcolor="#B8CCE4">����ʵ��֧��</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(sy_sj_zc,,-1)%></td>
    <td bgcolor="#B8CCE4">����֧�����</td>
    <td bgcolor="#DBE5F1"><%=FormatCurrency(sy_sj_zc-sy_jh_zc,,-1)%></td>
    <%
  	rs_sy.close
	set rs_sy=nothing
  %>
  </tr>
</table>
<br>
<form id="form1" name="form1" method="post" action="" onSubmit="return check_in(this);">
<table align="center" cellpadding="2" cellspacing="1" border="0" bgcolor="#5780B4" width="100%">
  <tr bgcolor="#B8CCE4">
    <th scope="col">�ƻ�֧��/����</th>
    <th scope="col">ʵ��֧��/����</th>
    <th scope="col">��ע</th>
    <th scope="col">֧��/��������</th>
    <th scope="col">֧��/������</th>
    <th scope="col">������</th>
    <th scope="col"><p>��д��</p></th>
  </tr>
  <tr align="center" bgcolor="#DBE5F1">
    <td><input name="JH_ZC" type="text" id="JH_ZC" value="0.00" size="10" /></td>
    <td><input name="SJ_ZC" type="text" id="SJ_ZC" value="0.00" size="10" /></td>
    <td><input name="BZ" type="text" id="BZ" size="10" /></td>
    <td><input name="ZC_RQ" type="text" id="ZC_RQ" value="<%=formatdatetime(now(),2)%>" size="10" />
	<script type="text/javascript">
		Calendar.setup({
			inputField     :    "ZC_RQ",   // id of the input field
			ifFormat       :    "%Y/%m/%d",       // format of the input field
			showsTime      :    false,
			timeFormat     :    "24",
			showOthers     :    true,
			eventName      :    "click",
			onUpdate       :    "ZC_RQ"
		});
	</script></td>
    <td><select name="ZCR" id="ZCR">
        <option value="СC">СC</option>
        <option value="СP">СP</option>
    </select></td>
    <td><select name="SYR" id="SYR">
        <option value="СC">СC</option>
        <option value="СP">СP</option>
        <option value="��ͥ">��ͥ</option>
        <option value="����">����</option>
    </select></td>
    <td><select name="TXR" id="TXR">
        <option value="СC">СC</option>
        <option value="СP">СP</option>
    </select></td>
  </tr>
  <tr bgcolor="#EFEFEF">
    <td colspan="9"><table cellpadding="2" cellspacing="1" border="0" bgcolor="#5780B4" width="100%">
        <tr bgcolor="#B8CCE4">
          <%
		set rs_lb=server.CreateObject("adodb.recordset")
		sql="select * from ֧����Ŀ��"
		rs_lb.open sql,conn
		while not rs_lb.eof
		%>
          <th scope="col"><%=rs_lb("XM_LB")%></th>
		  <% rs_lb.movenext
		  	wend %>
          </tr>
        <tr valign="top" bgcolor="#DBE5F1">
          <% rs_lb.movefirst 
		while not rs_lb.eof %>
          <td><%
			set rs_xm=server.CreateObject("adodb.recordset")
			sql="select * from ֧����Ŀ�б� where ZC_LB="&rs_lb("ID")
			rs_xm.open sql,conn
			while not rs_xm.eof
		  %>
		  <input type="radio" name="ZC_XM" value="<%=rs_xm("ID")%>" /><%=rs_xm("XM_MC")%><br />
		  <% rs_xm.movenext
				wend 
				rs_xm.close
				set rs_xm=nothing
				%></td>
          <% rs_lb.movenext
	  		wend 
			rs_lb.close
			set rs_lb=nothing
			%>
          </tr>
    </table></td></tr>
  <tr align="center" bgcolor="#B8CCE4">
    <td colspan="9">
      <input type="submit" name="Submit" value="�ύ" /> &nbsp;
      <input type="hidden" name="zc_add" value="true" />      
      &nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" value="����" /></td>
    </tr>
</table>
</form>
<form id="form2" name="form2" method="post">
<table align="center" cellpadding="2" cellspacing="1" border="0" bgcolor="#5780B4" width="100%">
  <tr bgcolor="#B8CCE4">
    <td>�������ݣ�<input name="more_sch" type="text" id="more_sch" size="20">
      ѡ�����ڣ�
        <input name="ksrq" type="text" id="ksrq" value="<%=formatdatetime(year(now())&"/"&month(now()),2)%>" size="10" maxlength="10" />
      -
      <input name="jsrq" type="text" id="jsrq" value="<%=formatdatetime(now(),2)%>" size="10" maxlength="10" />
      <input type="submit" name="view_list" value="�ύ" />
	  <script type="text/javascript">
		Calendar.setup({
			inputField     :    "ksrq",   // id of the input field
			ifFormat       :    "%Y/%m/%d",       // format of the input field
			showsTime      :    false,
			timeFormat     :    "24",
			showOthers     :    true,
			eventName      :    "click",
			onUpdate       :    "ksrq"
		});
		Calendar.setup({
			inputField     :    "jsrq",   // id of the input field
			ifFormat       :    "%Y/%m/%d",       // format of the input field
			showsTime      :    false,
			timeFormat     :    "24",
			showOthers     :    true,
			eventName      :    "click",
			onUpdate       :    "jsrq"
		});
	</script></td>
  </tr>
</table>
</form>
<table align="center" cellpadding="2" cellspacing="1" border="0" bgcolor="#5780B4" width="100%">
  <tr bgcolor="#B8CCE4">
    <th scope="col">ID</th>
    <th scope="col">�ƻ�֧��/����</th>
    <th scope="col">ʵ��֧��/����</th>
    <th scope="col">֧��/�������</th>
    <th scope="col">֧��/������Ŀ</th>
    <th scope="col">��ע</th>
    <th scope="col">֧��/������</th>
    <th scope="col">������</th>
    <th scope="col">��д��</th>
    <th scope="col">֧������</th>
  </tr>
  <% 
	if request.QueryString("xml_id")<>"" then 
		xml_id=trim(replace(request.QueryString("xml_id"),"'","''"))
		str_sql="and ֧����Ŀ��.id="&xml_id
	end if
	if request.QueryString("zc_xm")<>"" then 
		zc_xm=trim(replace(request.QueryString("zc_xm"),"'","''"))
		str_sql=" and zc_xm="&zc_xm
	end if
	if request.QueryString("zcr")<>"" then 
		zcr=trim(replace(request.QueryString("zcr"),"'","''"))
		str_sql=" and zcr='"&urldecoding(zcr)&"'"
	end if
	if request.QueryString("syr")<>"" then 
		syr=trim(replace(request.QueryString("syr"),"'","''"))
		str_sql=" and syr='"&urldecoding(syr)&"'"
	end if
	if request.QueryString("txr")<>"" then 
		txr=trim(replace(request.QueryString("txr"),"'","''"))
		str_sql=" and txr='"&urldecoding(txr)&"'"
	end if
	set rs=server.CreateObject("adodb.recordset")
	sql = "SELECT * FROM ��ͥ��֧���굥,֧����Ŀ�б�,֧����Ŀ�� where ��ͥ��֧���굥.ZC_XM=֧����Ŀ�б�.ID and ֧����Ŀ�б�.ZC_LB=֧����Ŀ��.ID "&str_sql&" and (ZC_RQ>=DateValue(Year(Now()) & '/' & Month(Now())) And ZC_RQ<=Now()) ORDER BY ��ͥ��֧���굥.DJ_RQ DESC" 
	If request.Form("view_list")="�ύ" Then 
		if request.form("ksrq")<>"" then ksrq=cdate(trim(replace(request.form("ksrq"),"'","''")))
		if request.form("jsrq")<>"" then jsrq=cdate(trim(replace(request.form("jsrq"),"'","''")))
		sql="SELECT * FROM ��ͥ��֧���굥,֧����Ŀ�б�,֧����Ŀ�� where ��ͥ��֧���굥.ZC_XM=֧����Ŀ�б�.ID and ֧����Ŀ�б�.ZC_LB=֧����Ŀ��.ID "&str_sql&" and (ZC_RQ>=DateValue('"&ksrq&"') And ZC_RQ<=DateValue('"&jsrq&"')) ORDER BY ��ͥ��֧���굥.DJ_RQ DESC" 
	End If
	response.Write(sql)
	rs.open sql,conn,1,1
	While NOT rs.EOF
  %>
  <tr align="center" bgcolor="<% if rs.AbsolutePosition=1 then %>#ffff99<% else %>#DBE5F1<% end if %>">
    <td><%=rs.AbsolutePosition%></td>
    <td><%=FormatCurrency((rs.Fields.Item("JH_ZC").Value),,-1)%></td>
    <td><%=FormatCurrency((rs.Fields.Item("SJ_ZC").Value),,-1)%></td>
    <td><a href="?xml_id=<%=rs("֧����Ŀ��.id")%>"><%=(rs.Fields.Item("XM_LB").Value)%></a></td>
    <td><a href="?zc_xm=<%=rs("zc_xm")%>"><%=(rs.Fields.Item("XM_MC").Value)%></a></td>
    <td><%=(rs.Fields.Item("BZ").Value)%></td>
    <td><a href="?zcr=<%=server.URLEncode(rs("zcr"))%>"><%=(rs.Fields.Item("ZCR").Value)%></a></td>
    <td><a href="?syr=<%=server.URLEncode(rs("syr"))%>"><%=(rs.Fields.Item("SYR").Value)%></a></td>
    <td><a href="?txr=<%=server.URLEncode(rs("txr"))%>"><%=(rs.Fields.Item("TXR").Value)%></a></td>
    <td><%=(rs.Fields.Item("ZC_RQ").Value)%></td>
  </tr>
  <% 
		jh_sum=jh_sum+rs.Fields.Item("JH_ZC").Value
		sj_sum=sj_sum+rs.Fields.Item("SJ_ZC").Value
  rs.MoveNext()
Wend
%>
  <tr align="center" bgcolor="#B8CCE4">
    <td>�ϼ�</td>
    <td><%=FormatCurrency(jh_sum,,-1)%></td>
    <td><%=FormatCurrency(sj_sum,,-1)%></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
</table>
<p>&nbsp;</p>

<%
rs.Close()
Set rs = Nothing
%>
</body>
</html>
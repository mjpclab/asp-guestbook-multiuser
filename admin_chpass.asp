<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� �޸Ĺ���Ա����</title>

	<script type="text/javascript">
	//<![CDATA[
	
	function checkpass(cobject)
	{
		if (cobject.ioldpass.value=="") {alert('������ԭ���롣'); cobject.ioldpass.focus(); return false;}
		if (cobject.inewpass1.value=="") {alert('�����������롣'); cobject.inewpass1.focus(); return false;}
		if (cobject.inewpass2.value=="") {alert('������ȷ�����롣'); cobject.inewpass2.focus(); return false;}
		if (cobject.inewpass1.value!=cobject.inewpass2.value) {alert('��������ȷ�����벻ͬ�����������롣'); cobject.inewpass1.focus(); return(false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	function checkqa(cobject)
	{
		if (cobject.ioldpass.value=="") {alert('���������롣'); cobject.ioldpass.focus(); return false;}
		if (cobject.question.value=="") {alert('�������һ��������⡣'); cobject.question.focus(); return false;}
		if (cobject.key.value=="") {alert('�������һ�����𰸡�'); cobject.key.focus(); return false;}
		cobject.submit1.disabled=true;
		return (true);
	}
	
	//]]>
	</script>
	
	<!-- #include file="style.asp" -->
</head>

<body<%=bodylimit%> onload="form4.ioldpass.focus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admintool.inc" -->

	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">�޸�����</td>
		</tr>
		<tr>
			<td class="wordscontent" style="text-align:center; padding:20px 0px;">
			<form method="post" action="admin_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
			��ԭ���룺<input type="password" name="ioldpass" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			�������룺<input type="password" name="inewpass1" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			ȷ�����룺<input type="password" name="inewpass2" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/><br/>
			<input value="��������" type="submit" name="submit1" />
			</form>
			</td>
		</tr>
	</table>

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype
	checkuser cn,rs,false

	rs.Open Replace(sql_adminchpass_question,"{0}",Request.QueryString("user")),cn,,,1
	%>
	
	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">�޸��һ���������/��</td>
		</tr>
		<tr>
			<td class="wordscontent" style="text-align:center; padding:20px 0px;">
			<form method="post" action="admin_saveqa.asp" onsubmit="return checkqa(this)" name="form5">
			<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
			���룺<input type="password" name="ioldpass" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			���⣺<input type="text" name="question" size="<%=SetInfoTextWidth%>" maxlength="32" value="<%=server.HTMLEncode("" & rs("question") & "")%>" /><br/>
			�𰸣�<input type="text" name="key" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/><br/>
			<input value="��������" type="submit" name="submit1" />
			</form>
			</td>
		</tr>
	</table>
	<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
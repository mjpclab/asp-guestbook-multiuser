<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires = -1

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,false

cn.Close
set rs=nothing
set cn=nothing
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster�������� �����û�����</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	//<![CDATA[
	
	function checkpass(cobject)
	{
		if (cobject.user.value=="") {alert('�������û�����'); cobject.user.focus(); return false;}
		if (cobject.inewpass1.value=="") {alert('�����������롣'); cobject.inewpass1.focus(); return false;}
		if (cobject.inewpass2.value=="") {alert('������ȷ�����롣'); cobject.inewpass2.focus(); return false;}
		if (cobject.inewpass1.value!=cobject.inewpass2.value) {alert('��������ȷ�����벻ͬ�����������롣'); cobject.inewpass1.focus(); return false;}
		cobject.submit1.disabled=true;
		return true;
	}
	
	//]]>
	</script>
	
</head>

<body onload="form4.inewpass1.focus();">

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->

	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">�����û�����</td>
		</tr>
		<tr>
			<td style="width:100%; text-align:center; color:<%=TableContentColor%>; background-color:<%=TableContentBGC%>;">
			<br/>
				<form method="post" action="web_saveuserpass.asp" onsubmit="return checkpass(this)" name="form4">
				���û�����<input type="text" name="user" size="<%=SetInfoTextWidth%>" maxlength="32" value="<%=Request.QueryString("user")%>" /><br/>
				�������룺<input type="password" name="inewpass1" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
				ȷ�����룺<input type="password" name="inewpass2" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/><br/>
				<input value="��������" type="submit" name="submit1" />
				</form>
			<br/>
		</td></tr>
	</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
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
	<!-- #include file="inc_web_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function checkpass(cobject)
	{
		if (cobject.user.value.length===0) {alert('�������û�����'); cobject.user.focus(); return false;}
		if (cobject.inewpass1.value.length===0) {alert('�����������롣'); cobject.inewpass1.focus(); return false;}
		if (cobject.inewpass2.value.length===0) {alert('������ȷ�����롣'); cobject.inewpass2.focus(); return false;}
		if (cobject.inewpass1.value!==cobject.inewpass2.value) {alert('��������ȷ�����벻ͬ�����������롣'); cobject.inewpass1.focus(); return false;}
		cobject.submit1.disabled=true;
		return true;
	}
	</script>
	
</head>

<body onload="form4.inewpass1.focus();">

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->

	<div class="region form-region">
		<h3 class="title">�����û�����</h3>
		<div class="content">
			<form method="post" action="web_saveuserpass.asp" onsubmit="return checkpass(this)" name="form4">
			<div class="field">
				<span class="label">�û�����</span>
				<span class="value"><input type="text" name="user" maxlength="32" value="<%=ruser%>" /></span>
			</div>
			<div class="field">
				<span class="label">�����룺</span>
				<span class="value"><input type="password" name="inewpass1" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">ȷ�����룺</span>
				<span class="value"><input type="password" name="inewpass2" maxlength="32" /></span>
			</div>
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
            </form>
		</div>
	</div>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
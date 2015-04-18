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
	<title><%=web_BookName%> Webmaster管理中心 重设用户密码</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function checkpass(cobject)
	{
		if (cobject.user.value.length===0) {alert('请输入用户名。'); cobject.user.focus(); return false;}
		if (cobject.inewpass1.value.length===0) {alert('请输入新密码。'); cobject.inewpass1.focus(); return false;}
		if (cobject.inewpass2.value.length===0) {alert('请输入确认密码。'); cobject.inewpass2.focus(); return false;}
		if (cobject.inewpass1.value!==cobject.inewpass2.value) {alert('新密码与确认密码不同，请重新输入。'); cobject.inewpass1.focus(); return false;}
		cobject.submit1.disabled=true;
		return true;
	}
	</script>
	
</head>

<body onload="form4.inewpass1.focus();">

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->

	<div class="region form-region">
		<h3 class="title">重设用户密码</h3>
		<div class="content">
			<form method="post" action="web_saveuserpass.asp" onsubmit="return checkpass(this)" name="form4">
			<div class="field">
				<span class="label">用户名：</span>
				<span class="value"><input type="text" name="user" maxlength="32" value="<%=ruser%>" /></span>
			</div>
			<div class="field">
				<span class="label">新密码：</span>
				<span class="value"><input type="password" name="inewpass1" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">确认密码：</span>
				<span class="value"><input type="password" name="inewpass2" maxlength="32" /></span>
			</div>
			<div class="command"><input value="更新数据" type="submit" name="submit1" /></div>
            </form>
		</div>
	</div>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
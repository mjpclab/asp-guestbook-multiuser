<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster管理中心 修改管理员密码</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function checkpass(cobject)
	{
		if (cobject.ioldpass.value.length===0) {alert('请输入原密码。'); cobject.ioldpass.focus(); return(false);}
		if (cobject.inewpass1.value.length===0) {alert('请输入新密码。'); cobject.inewpass1.focus(); return(false);}
		if (cobject.inewpass2.value.length===0) {alert('请输入确认密码。'); cobject.inewpass2.focus(); return(false);}
		if (cobject.inewpass1.value!==cobject.inewpass2.value) {alert('新密码与确认密码不同，请重新输入。'); cobject.inewpass1.focus(); return(false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	</script>
	
</head>

<body onload="form4.ioldpass.focus();">

<div id="outerborder" class="outerborder">

	<!-- #include file="include/template/web_admin_header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<div class="region form-region region-longtext">
		<h3 class="title">修改密码</h3>
		<div class="content">
			<form method="post" action="web_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			<div class="field">
				<span class="label">原密码：</span>
				<span class="value"><input type="password" name="ioldpass" maxlength="32" /></span>
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

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
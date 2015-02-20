<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster管理中心 修改管理员密码</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	function checkpass(cobject)
	{
		if (cobject.ioldpass.value=="") {alert('请输入原密码。'); cobject.ioldpass.focus(); return(false);}
		if (cobject.inewpass1.value=="") {alert('请输入新密码。'); cobject.inewpass1.focus(); return(false);}
		if (cobject.inewpass2.value=="") {alert('请输入确认密码。'); cobject.inewpass2.focus(); return(false);}
		if (cobject.inewpass1.value!=cobject.inewpass2.value) {alert('新密码与确认密码不同，请重新输入。'); cobject.inewpass1.focus(); return(false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	</script>
	
</head>

<body onload="form4.ioldpass.focus();">

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admintool.inc" -->

	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">修改密码</td>
		</tr>
		<tr>
		<td class="wordscontent" style="text-align:center; padding:20px 0px;">
			<form method="post" action="web_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			　原密码：<input type="password" name="ioldpass" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			　新密码：<input type="password" name="inewpass1" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			确认密码：<input type="password" name="inewpass2" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/><br/>
			<input value="更新数据" type="submit" name="submit1" />
			</form>
		</td>
		</tr>
	</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
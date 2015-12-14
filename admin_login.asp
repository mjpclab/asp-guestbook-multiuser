<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif StatusLogin=false then
	Call WebErrorPage(3)
	Response.End
end if

if VcodeCount>0 then Session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 管理员登录</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
	
	<script type="text/javascript">
	function submitCheck(obj)
	{
		if(obj.user.value.length===0)
		{
			alert('请输入用户名。');
			obj.user.focus();
			return false;
		}
		else if(obj.iadminpass.value.length===0)
		{
			alert('请输入密码。');
			obj.iadminpass.focus();
			return false;
		}
		else if(obj.ivcode && obj.ivcode.value.length===0)
		{
			alert('请输入验证码。');
			obj.ivcode.focus();
			return false;
		}
		else
		{
			obj.submit1.disabled=true;
			return true;
		}
	}
	</script>
</head>

<body onload="if(form5.iadminpass.value.length===0)form5.iadminpass.focus()">

<div class="region form-region narrow-region">
	<h3 class="title">用户登录</h3>
	<div class="content">
		<form method="post" action="login_verify.asp" name="form5" onsubmit="return submitCheck(this);">
			<div class="field">
				<span class="label">用户名：</span>
				<span class="value"><input type="text" name="user" maxlength="32" value="<%=ruser%>" /></span>
			</div>
			<div class="field">
				<span class="label">密码：</span>
				<span class="value"><input type="password" name="iadminpass" maxlength="32" autofocus="autofocus" /></span>
			</div>
			<%if VcodeCount>0 then%>
			<div class="field">
				<span class="label">验证码：</span>
				<span class="value"><input type="text" name="ivcode" autocomplete="off" /><img id="captcha" class="captcha" src="show_vcode.asp?user=<%=ruser%>&t=0"/></span>
			</div>
			<%end if%>
			<div class="command">
				<input type="submit" value="登录" name="submit1" />
			</div>
			</form>
	</div>
</div>
<script type="text/javascript">
	<!-- #include file="asset/js/refresh-captcha.js" -->
</script>
</body>
</html>
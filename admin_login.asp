<!-- #include file="loadconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusLogin=false then
	Response.Redirect "web_err.asp?number=3"
	Response.End
end if

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 管理员登录</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
	
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
				<span class="value"><input type="text" name="user" maxlength="32" value="<%=Request.QueryString("user")%>" /></span>
			</div>
			<div class="field">
				<span class="label">密码：</span>
				<span class="value"><input type="password" name="iadminpass" maxlength="32" autofocus="autofocus" /></span>
			</div>
			<%if VcodeCount>0 then%>
			<div class="field">
				<span class="label">验证码：</span>
				<span class="value"><input type="text" name="ivcode" autocomplete="off" /><img class="captcha" src="show_vcode.asp?user=<%=Request.QueryString("user")%>"/></span>
			</div>
			<%end if%>
			<div class="command">
				<input type="submit" value="登录" name="submit1" />
			</div>
			</form>
	</div>
</div>

</body>
</html>
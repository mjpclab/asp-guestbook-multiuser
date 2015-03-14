<!-- #include file="webconfig.asp" -->
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
	<!-- #include file="css/style.asp" -->
</head>

<body onload="if(form5.user.value.length===0)form5.user.focus();else form5.iadminpass.focus()">

<div class="region form-region narrow-region">
	<h3 class="title">用户登录</h3>
	<div class="content">
		<form method="post" action="login_verify.asp" name="form5" onsubmit="return submitCheck(this);">
			<div class="field">
				<span class="label">用户名：</span>
				<span class="value"><input type="text" name="user" maxlength="32" value="<%=Request.QueryString("user")%>" <%if Request.QueryString("user")="" then Response.Write "autofocus='autofocus'"%>/></span>
			</div>
			<div class="field">
				<span class="label">密码：</span>
				<span class="value"><input type="password" name="iadminpass" maxlength="32" <%if Request.QueryString("user")<>"" then Response.Write "autofocus='autofocus'"%>/></span>
			</div>
			<%if VcodeCount>0 then%>
			<div class="field">
				<span class="label">验证码：</span>
				<span class="value"><input type="text" name="ivcode" size="10" autocomplete="off" /><img class="captcha" src="web_show_vcode.asp"/></span>
			</div>
			<%end if%>
			<div class="command">
				<input value="登录" type="submit" name="submit1" />
			</div>
			</form>
	</div>
</div>

</body>
</html>
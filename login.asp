<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/web.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif Not StatusLogin then
	Call WebErrorPage(3)
	Response.End
end if

if VcodeCount>0 then Session(InstanceName & "_vcode")=getvcode(VcodeCount)
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> 留言本 管理员登录</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body onload="if(form5.user.value.length===0)form5.user.focus();else form5.iadminpass.focus()">
<%Call WebInitHeaderData("","管理员登录","","")%><!-- #include file="include/template/header.inc" -->

<div id="mainborder" class="mainborder narrow-mainborder">
<div class="region form-region narrow-region">
	<h3 class="title">管理员登录</h3>
	<div class="content">
		<form method="post" action="login_verify.asp" name="form5" onsubmit="return submitCheck(this);">
			<div class="field">
				<span class="label">用户名</span>
				<span class="value"><input type="text" name="user" maxlength="32" value="<%=ruser%>" <%if ruser="" then Response.Write "autofocus='autofocus'"%>/></span>
			</div>
			<div class="field">
				<span class="label">密码</span>
				<span class="value"><input type="password" name="iadminpass" maxlength="32" <%if ruser<>"" then Response.Write "autofocus='autofocus'"%>/></span>
			</div>
			<%if VcodeCount>0 then%>
			<div class="field">
				<span class="label">验证码</span>
				<span class="value"><input type="text" name="ivcode" size="10" autocomplete="off" /><img id="captcha" class="captcha" src="show_vcode.asp?t=0"/></span>
			</div>
			<%end if%>
			<div class="command">
				<input value="登录" type="submit" name="submit1" />
			</div>
			</form>
	</div>
</div>
</div>
<script type="text/javascript">
	<!-- #include file="asset/js/refresh-captcha.js" -->
</script>
</body>
</html>
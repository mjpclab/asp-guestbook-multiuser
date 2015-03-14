<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusUnreg=false then
	Response.Redirect "web_err.asp?number=5"
	Response.End
end if

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> 自删帐号</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<!-- #include file="css/style.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.user.value.length===0) {alert('请输入用户名。'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)===false) {alert('用户名中只能包含英文字母、数字和下划线。'); frm.user.focus(); return false;}
		else if (frm.pass1.value.length===0) {alert('请输入密码。'); frm.pass1.focus(); return false;}
		else if (frm.vcode && frm.vcode.value.length===0) {alert('请输入验证码。'); frm.vcode.focus(); return false;}
		else return confirm('删除用户名将同时清除全部留言、回复及公告，确实要继续吗？');
	}
	</script>
	
</head>

<body onload="unregform.user.focus();">

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; 自删帐号</div>
</div>
<%
set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="func_web.inc" -->

<div class="region form-region">
	<h3 class="title">自删帐号</h3>
	<div class="content">
		<form name="unregform" method="post" action="submitunreg.asp" onsubmit="return submitcheck(this)">

		<div class="field">
			<span class="label">用户名：</span>
			<span class="value"><input type="text" name="user" maxlength="32" size="32" title="只能包含英文、数字和下划线，最大32位" /></span>
		</div>
		<div class="field">
			<span class="label">密码：</span>
			<span class="value"><input type="password" name="pass1" maxlength="32" size="32" /></span>
		</div>
		<%if VcodeCount>0 then%>
		<div class="field">
			<span class="label">验证码：</span>
			<span class="value"><input type="text" name="vcode" size="10" /><img class="captcha" src="web_show_vcode.asp"/></span>
		</div>
		<%end if%>
		<div class="command"><input type="submit" value="提交" />　　<input type="reset" value="重写" /></div>

		</form>
	</div>
</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
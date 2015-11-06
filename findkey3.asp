<!-- #include file="webconfig.asp" -->
<!-- #include file="md5.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusFindkey=false then
	Response.Redirect "web_err.asp?number=2"
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	Call MessagePage("验证码错误。","findkey.asp")
	Response.End
elseif VcodeCount>0 then
	session("vcode")=getvcode(VcodeCount)
else
	session("vcode")=""
end if

'===============================合式验证
dim re
set re=new RegExp
re.Pattern="^\w+$"

ruser=Request.Form("user")
if ruser="" then
	Call MessagePage("用户名不能为空。","findkey.asp")
	Response.End
elseif re.Test(ruser)=false then
	Call MessagePage("用户名中只能包含英文字母、数字和下划线。","findkey.asp")
	Response.End
elseif Request.Form("key")="" then
	Call MessagePage("找回密码答案不能为空。","findkey.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

'=============================存在性验证
rs.Open Replace(sql_findkey3_checkuser,"{0}",ruser),cn,,,1
if rs.EOF then
	Call MessagePage("用户名不存在。","findkey.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if

if md5(Request.Form("key"),32)<>rs("key") then
	Call MessagePage("找回密码答案不正确。","findkey.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : set rs=nothing
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> 找回密码步骤3</title>
	<!-- #include file="inc_stylesheet.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.pass1.value.length===0) {alert('请输入密码。'); frm.pass1.focus(); return false;}
		else if (frm.pass2.value.length===0) {alert('请输入确认密码。'); frm.pass2.focus(); return false;}
		else if (frm.pass1.value!==frm.pass2.value) {alert('密码不一致，请检查。'); frm.pass2.select(); return false;}
		else if (frm.vcode && frm.vcode.value.length===0) {alert('请输入验证码。'); frm.vcode.focus(); return false;}
		else {frm.submit1.disabled=true; return true;}
	}
	</script>
</head>

<body onload="findform3.pass1.focus();">

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">找回密码</a> &gt;&gt; 步骤3</div>
</div>

<%
dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="func_web.inc" -->

<div class="region form-region">
	<h3 class="title">找回密码 步骤3</h3>
	<div class="content">
		<form name="findform3" method="post" action="findkey4.asp" onsubmit="return submitcheck(this)">
			<input type="hidden" name="user" value="<%=ruser%>" />
			<input type="hidden" name="key" value="<%=Request.Form("key")%>" />

			<h4>请重新设置您的密码</h4>
			<div class="field">
				<span class="label">密码：</span>
				<span class="value"><input type="password" name="pass1" maxlength="32" size="32" /></span>
			</div>
			<div class="field">
				<span class="label">确认密码：</span>
				<span class="value"><input type="password" name="pass2" maxlength="32" size="32" /></span>
			</div>
			<%if VcodeCount>0 then%>
			<div class="field">
				<span class="label">验证码：</span>
				<span class="value"><input type="text" name="vcode" size="10" autocomplete="off" /><img id="captcha" class="captcha" src="web_show_vcode.asp?t=0"/></span>
			</div>
			<%end if%>
			<div class="command"><input type="submit" name="submit1" value="完成重设" />　　<input type="reset" value="重新填写" /></div>
		</form>
	</div>
</div>

</div>
<!-- #include file="bottom.asp" -->
<script type="text/javascript">
	<!-- #include file="js/refresh-captcha.js" -->
</script>
</body>
</html>
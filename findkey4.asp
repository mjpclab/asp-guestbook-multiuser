<!-- #include file="webconfig.asp" -->
<!-- #include file="include/md5.asp" -->
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
	Call MessagePage("验证码错误","findkey.asp")
	Response.End
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
elseif Request.Form("pass1")="" then
	Call MessagePage("密码不能为空。","findkey.asp")
	Response.End
elseif Request.Form("pass2")="" then
	Call MessagePage("确认密码不能为空。","findkey.asp")
	Response.End
elseif Request.Form("pass1")<>Request.Form("pass2") then
	Call MessagePage("密码不一致，请检查。","findkey.asp")
	Response.End
elseif len(Request.Form("pass1"))>32 then
	Call MessagePage("密码长度不能超过32字。","findkey.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

'=============================存在性验证
rs.Open Replace(sql_findkey4_checkuser,"{0}",ruser),cn,,,1
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
cn.Execute Replace(Replace(sql_findkey4_resetpass,"{0}",md5(Request.Form("pass1"),32)),"{1}",ruser),,1

%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=web_BookName%> 找回密码完成</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">找回密码</a> &gt;&gt; 完成</div>
</div>

<%
dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="include/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/web_guest_func.inc" -->

<div class="region form-region">
	<h3 class="title">找回密码 步骤3</h3>
	<div class="content">
		<h4>已重设密码，请点击下面的链接登录：</h4>
		<p><a href="admin_login.asp?user=<%=ruser%>">用户登录</a></p>
	</div>
</div>

</div>

<!-- #include file="include/footer.inc" -->
</body>
</html>
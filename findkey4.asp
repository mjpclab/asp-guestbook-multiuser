<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/sql/web_findkey4.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="tips.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif Not StatusFindkey then
	Call WebErrorPage(2)
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>Session("vcode") or Session("vcode")="") then
	Session("vcode")=""
	Call TipsPage("验证码错误","findkey.asp")
	Response.End
else
	Session("vcode")=""
end if

'===============================合式验证
dim re
set re=new RegExp
re.Pattern="^\w+$"

ruser=Request.Form("user")
if ruser="" then
	Call TipsPage("用户名不能为空。","findkey.asp")
	Response.End
elseif re.Test(ruser)=false then
	Call TipsPage("用户名中只能包含英文字母、数字和下划线。","findkey.asp")
	Response.End
elseif Request.Form("key")="" then
	Call TipsPage("找回密码答案不能为空。","findkey.asp")
	Response.End
elseif Request.Form("pass1")="" then
	Call TipsPage("密码不能为空。","findkey.asp")
	Response.End
elseif Request.Form("pass2")="" then
	Call TipsPage("确认密码不能为空。","findkey.asp")
	Response.End
elseif Request.Form("pass1")<>Request.Form("pass2") then
	Call TipsPage("密码不一致，请检查。","findkey.asp")
	Response.End
elseif len(Request.Form("pass1"))>32 then
	Call TipsPage("密码长度不能超过32字。","findkey.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

'=============================存在性验证
rs.Open Replace(sql_findkey4_checkuser,"{0}",ruser),cn,,,1
if rs.EOF then
	Call TipsPage("用户名不存在。","findkey.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if

if md5(Request.Form("key"),32)<>rs("key") then
	Call TipsPage("找回密码答案不正确。","findkey.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : set rs=nothing
cn.Execute Replace(Replace(sql_findkey4_resetpass,"{0}",md5(Request.Form("pass1"),32)),"{1}",ruser),,1

%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> 找回密码完成</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">找回密码</a> &gt;&gt; 完成</div>
</div>
<div id="mainborder" class="mainborder">
<%
dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="include/template/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/template/web_guest_func.inc" -->

<div class="region form-region">
	<h3 class="title">找回密码 步骤3</h3>
	<div class="content">
		<h4>已重设密码，请点击下面的链接登录：</h4>
		<p><a href="admin_login.asp?user=<%=ruser%>">用户登录</a></p>
	</div>
</div>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
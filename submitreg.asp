<!-- #include file="webconfig.asp" -->
<!-- #include file="common2.asp" -->
<!-- #include file="md5.asp" -->
<%

Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusReg=false then
	Response.Redirect "web_err.asp?number=1"
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	Call MessagePage("验证码错误。","reg.asp")
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
	Call MessagePage("用户名不能为空。","reg.asp")
	Response.End
elseif re.Test(ruser)=false then
	Call MessagePage("用户名中只能包含英文字母、数字和下划线。","reg.asp")
	Response.End
elseif len(ruser)>32 then
	Call MessagePage("用户名长度不能超过32字。","reg.asp")
	Response.End
elseif Request.Form("pass1")="" then
	Call MessagePage("密码不能为空。","reg.asp")
	Response.End
elseif Request.Form("pass2")="" then
	Call MessagePage("确认密码不能为空。","reg.asp")
	Response.End
elseif Request.Form("pass1")<>Request.Form("pass2") then
	Call MessagePage("密码不一致，请检查。","reg.asp")
	Response.End
elseif len(Request.Form("pass1"))>32 then
	Call MessagePage("密码长度不能超过32字。","reg.asp")
	Response.End
elseif Request.Form("question")="" then
	Call MessagePage("找回密码问题不能为空。","reg.asp")
	Response.End
elseif Request.Form("key")="" then
	Call MessagePage("找回密码答案不能为空。","reg.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

'=============================存在性验证
rs.Open Replace(sql_submitreg_checkuser,"{0}",ruser),cn,,,1
if not rs.EOF then
	Call MessagePage("用户名已存在，请重新输入。","reg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close

dim tnow
cn.BeginTrans
tnow=now()
cn.Execute Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_submitreg_init1,"{0}",ruser),"{1}",md5(Request.Form("pass1"),32)),"{2}",FilterQuote(server.HTMLEncode(Request.Form("nick")))),"{3}",0),"{4}",tnow),"{5}",tnow),"{6}",FilterQuote(Request.Form("question"))),"{7}",md5(Request.Form("key"),32)),"{8}",""),,1
cn.Execute Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_submitreg_init2,"{0}",ruser),"{1}",143),"{2}",6),"{3}",6),"{4}",20),"{5}",34),"{6}",132),"{7}",132),"{8}",111),"{9}",1270),"{10}",styleid),,1
cn.Execute Replace(sql_submitreg_init3,"{0}",ruser),,1
cn.CommitTrans

set rs=nothing

dim gbookaddr
gbookaddr=geturlpath & "index.asp?user=" & ruser
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> 创建成功</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; <a href="reg.asp" style="color:<%=TitleColor%>">申请留言本</a> &gt;&gt; 申请成功</div>
</div>

<%
dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="func_web.inc" -->

<p class="centertext">创建成功！请记下您的留言本首页地址：<br/><%=gbookaddr%><br/>&gt;<a href="<%=gbookaddr%>">转到该页→</a></p>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/sql/web_submitreg.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="tips.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif Not StatusReg then
	Call WebErrorPage(1)
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>Session(InstanceName & "_vcode") or Session(InstanceName & "_vcode")="") then
	Session(InstanceName & "_vcode")=""
	Call TipsPage("验证码错误。","reg.asp")
	Response.End
else
	Session(InstanceName & "_vcode")=""
end if
'===============================合式验证
dim re
set re=new RegExp
re.Pattern="^\w+$"

ruser=Request.Form("user")
if ruser="" then
	Call TipsPage("用户名不能为空。","reg.asp")
	Response.End
elseif re.Test(ruser)=false then
	Call TipsPage("用户名中只能包含英文字母、数字和下划线。","reg.asp")
	Response.End
elseif len(ruser)>32 then
	Call TipsPage("用户名长度不能超过32字。","reg.asp")
	Response.End
elseif Request.Form("pass1")="" then
	Call TipsPage("密码不能为空。","reg.asp")
	Response.End
elseif Request.Form("pass2")="" then
	Call TipsPage("确认密码不能为空。","reg.asp")
	Response.End
elseif Request.Form("pass1")<>Request.Form("pass2") then
	Call TipsPage("密码不一致，请检查。","reg.asp")
	Response.End
elseif len(Request.Form("pass1"))>32 then
	Call TipsPage("密码长度不能超过32字。","reg.asp")
	Response.End
elseif Request.Form("question")="" then
	Call TipsPage("找回密码问题不能为空。","reg.asp")
	Response.End
elseif Request.Form("key")="" then
	Call TipsPage("找回密码答案不能为空。","reg.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

'=============================存在性验证
rs.Open Replace(sql_submitreg_checkuser,"{0}",ruser),cn,,,1
if not rs.EOF then
	Call TipsPage("用户名已存在，请重新输入。","reg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close

dim tnow
tnow=ServerTimeToUTC(now())
cn.BeginTrans
cn.Execute Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_submitreg_init1,"{0}",ruser),"{1}",md5(Request.Form("pass1"),32)),"{2}",FilterQuote(HtmlEncode(Request.Form("nick")))),"{3}",0),"{4}",tnow),"{5}",tnow),"{6}",FilterQuote(Request.Form("question"))),"{7}",md5(Request.Form("key"),32)),"{8}",""),,1
cn.Execute Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_submitreg_init2,"{0}",ruser),"{1}",143),"{2}",6),"{3}",6),"{4}",20),"{5}",34),"{6}",132),"{7}",132),"{8}",DisplayTimezoneOffset),"{9}",107),"{10}",1270),"{11}",CssFontFamily),"{12}",CssFontSize),"{13}",CssLineHeight),"{14}",styleid),,1
cn.Execute Replace(sql_submitreg_init3,"{0}",ruser),,1
cn.CommitTrans

set rs=nothing

dim gbookaddr
gbookaddr=geturlpath & "index.asp?user=" & ruser
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> 创建成功</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<%Call WebInitHeaderData("reg.asp","申请留言本","","申请成功")%><!-- #include file="include/template/header.inc" -->
<div id="mainborder" class="mainborder">
<%
dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="include/template/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/template/web_guest_func.inc" -->

<p class="centertext">创建成功！请记下您的留言本首页地址：<br/><%=gbookaddr%><br/>&gt;<a href="<%=gbookaddr%>">转到该页→</a></p>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
<!-- #include file="webconfig.asp" -->
<!-- #include file="md5.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
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

if Request.Form("user")="" then
	Call MessagePage("用户名不能为空。","findkey.asp")
	Response.End
elseif re.Test(Request.Form("user"))=false then
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
rs.Open Replace(sql_findkey4_checkuser,"{0}",Request.Form("user")),cn,,,1
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
cn.Execute Replace(Replace(sql_findkey4_resetpass,"{0}",md5(Request.Form("pass1"),32)),"{1}",Request.Form("user")),,1

%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=web_BookName%> 找回密码完成</title>

	<!-- #include file="style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">
	<table cellpadding="2" style="width:100%; border-width:0px; border-collapse:collapse;">
		<tr style="height:60px;">
			<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);" nowrap="nowrap">
			<%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">找回密码</a> &gt;&gt; 完成
			</td>
		</tr>		
		<tr>
			<td>
				<%			
				dim sys_bul_flag
				sys_bul_flag=32
				%>
				<!-- #include file="sysbulletin.inc" -->
				<%cn.Close : set cn=nothing%>
			</td>
		</tr>
		<tr>
			<td style="width:100%">
			<!-- #include file="func_web.inc" -->
			</td>
		</tr>
			<td style="width:100%; text-align:center; color:<%=TableContentColor%>;">
			<br/>
			已重设密码，请点击下面的链接登录：<br/><br/>
			<a href="admin_login.asp?user=<%=Request.Form("user")%>">用户登录</a><br/><br/>
			</td>
		<tr>
		</tr>
	</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
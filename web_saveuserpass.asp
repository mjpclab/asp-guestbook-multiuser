<!-- #include file="webconfig.asp" -->
<link rel="stylesheet" type="text/css" href="style.css"/>
<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
<link rel="stylesheet" type="text/css" href="web_adminstyle.css"/>
<!-- #include file="style.asp" -->
<!-- #include file="adminstyle.asp" -->
<!-- #include file="web_adminstyle.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="md5.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

checkuser cn,rs,true

if request.Form("inewpass1")<> request.Form("inewpass2") then
	Call MessagePage("新密码不一致，请检查。","web_chuserpass.asp?user=" & Request.Form("user"))
elseif request.Form("inewpass1")="" then
	Call MessagePage("密码不能为空，请重新输入。","web_chuserpass.asp?user=" & Request.Form("user"))
else
	cn.Execute Replace(Replace(sql_websaveuserpass,"{0}",md5(Request.Form("inewpass1"),32)),"{1}",Request.Form("user")),,1
	cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "web_userinfo.asp?user=" & Request.Form("user")
end if
cn.Close : set rs=nothing : set cn=nothing
%>
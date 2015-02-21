<!-- #include file="webconfig.asp" -->
<link rel="stylesheet" type="text/css" href="style.css"/>
<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
<link rel="stylesheet" type="text/css" href="web_adminstyle.css"/>
<!-- #include file="style.asp" -->
<!-- #include file="adminstyle.asp" -->
<!-- #include file="web_adminstyle.asp" -->
<!-- #include file="style.asp" -->
<!-- #include file="md5.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if VcodeCount>0 and (Request.Form("ivcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	Call MessagePage("验证码错误。","web_login.asp")
	Response.End
else
	session("vcode")=""
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
rs.Open sql_web_login_verify,cn,0,1,1

session.Contents(InstanceName & "_webpass")=md5(request("iadminpass"),32)
if session.Contents(InstanceName & "_webpass")=rs(0) then
	session.Timeout=cint(AdminTimeOut)
	
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "web_admin.asp"
else
	Call MessagePage("密码不正确。","web_login.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
end if
%>
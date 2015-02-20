<!-- #include file="loadconfig.asp" -->
<link rel="stylesheet" type="text/css" href="style.css"/>
<!-- #include file="style.asp" -->
<!-- #include file="md5.asp" -->

<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusLogin=false then
	Response.Redirect "web_err.asp?number=3"
	Response.End
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
checkuser cn,rs,true

if StatusStatistics then call addstat("login")

if VcodeCount>0 and (Request.Form("ivcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	if StatusStatistics then call addstat("loginfailed")
	Call MessagePage("验证码错误。","admin_login.asp?user=" &request("user"))

	set rs=nothing
	set cn=nothing
	Response.End
else
	session("vcode")=""
end if

rs.Open Replace(sql_loginverify,"{0}",Request.Form("user")),cn,0,1,1

session.Contents(InstanceName & "_adminpass_" & Request.Form("user"))=md5(request("iadminpass"),32)
if session.Contents(InstanceName & "_adminpass_"& Request.Form("user"))=rs(0) then
	session.Timeout=cint(AdminTimeOut)

	cn.Execute Replace(Replace(sql_updatelastlogin,"{0}",now()),"{1}",Request.Form("user")),,1
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "admin.asp?user=" &Request.Form("user")
else
	if StatusStatistics then call addstat("loginfailed")
	Call MessagePage("密码不正确。","admin_login.asp?user=" &request("user"))

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>
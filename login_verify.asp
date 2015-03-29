<!-- #include file="loadconfig.asp" -->
<link rel="stylesheet" type="text/css" href="css/style.css"/>
<!-- #include file="css/style.asp" -->
<!-- #include file="md5.asp" -->

<%
Response.Expires=-1
if web_checkIsBannedIP then
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
	Call MessagePage("验证码错误。","admin_login.asp?user=" &ruser)

	set rs=nothing
	set cn=nothing
	Response.End
else
	session("vcode")=""
end if

rs.Open Replace(sql_loginverify,"{0}",adminid),cn,0,1,1

session.Contents(InstanceName & "_adminpass_" & ruser)=md5(request("iadminpass"),32)
if session.Contents(InstanceName & "_adminpass_"& ruser)=rs(0) then
	session.Timeout=cint(AdminTimeOut)

	cn.Execute Replace(Replace(sql_updatelastlogin,"{0}",now()),"{1}",adminid),,1
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "admin.asp?user=" &ruser
else
	if StatusStatistics then call addstat("loginfailed")
	Call MessagePage("密码不正确。","admin_login.asp?user=" &ruser)

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>
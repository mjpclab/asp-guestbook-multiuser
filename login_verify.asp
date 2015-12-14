<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="tips.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif StatusLogin=false then
	Call WebErrorPage(3)
	Response.End
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

Call CreateConn(cn)

if StatusStatistics then call addstat("login")

if VcodeCount>0 and (Request.Form("ivcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	if StatusStatistics then call addstat("loginfailed")
	Call TipsPage("验证码错误。","admin_login.asp?user=" &ruser)

	set rs=nothing
	set cn=nothing
	Response.End
else
	session("vcode")=""
end if

rs.Open Replace(sql_adminverify,"{0}",adminid),cn,0,1,1

session.Contents(InstanceName & "_adminpass_" & ruser)=md5(request("iadminpass"),32)
if session.Contents(InstanceName & "_adminpass_"& ruser)=rs(0) then
	session.Timeout=cint(AdminTimeOut)

	cn.Execute Replace(Replace(sql_updatelastlogin,"{0}",now()),"{1}",adminid),,1
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "admin.asp?user=" &ruser
else
	if StatusStatistics then call addstat("loginfailed")
	Call TipsPage("密码不正确。","admin_login.asp?user=" &ruser)

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>
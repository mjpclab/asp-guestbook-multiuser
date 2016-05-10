<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/sql/web_admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="tips.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if VcodeCount>0 and (Request.Form("ivcode")<>Session(InstanceName & "_vcode") or Session(InstanceName & "_vcode")="") then
	Session(InstanceName & "_vcode")=""
	Call TipsPage("验证码错误。","web_login.asp")
	Response.End
else
	Session(InstanceName & "_vcode")=""
end if

Dim cn,rs,refPass
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

Call CreateConn(cn)
rs.Open sql_web_admin_verify,cn,0,1,1
if Not rs.EOF then
	refPass=rs.Fields(0)
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing

Session(InstanceName & "_webpass")=md5(request("iadminpass"),32)
if Session(InstanceName & "_webpass")=refPass then
	session.Timeout=AdminTimeOut
	Response.Redirect "web_admin.asp"
else
	Call TipsPage("密码不正确。","web_login.asp")
end if
%>
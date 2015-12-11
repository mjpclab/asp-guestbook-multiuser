<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/sql/web_admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="tips.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if VcodeCount>0 and (Request.Form("ivcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	Call TipsPage("验证码错误。","web_login.asp")
	Response.End
else
	session("vcode")=""
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
rs.Open sql_web_admin_verify,cn,0,1,1

session.Contents(InstanceName & "_webpass")=md5(request("iadminpass"),32)
if session.Contents(InstanceName & "_webpass")=rs(0) then
	session.Timeout=cint(AdminTimeOut)
	
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "web_admin.asp"
else
	Call TipsPage("密码不正确。","web_login.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
end if
%>
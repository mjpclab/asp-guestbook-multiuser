<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_mpassaudit.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
if Request.Form("seltodel")="" then
	if isnumeric(Request.Form("tpage")) and Request.Form("tpage")<>"" then
		Response.Redirect "admin.asp?user=" &ruser& "&page=" & Request.Form("tpage")
	else
		Response.Redirect "admin.asp?user=" &ruser
	end if
end if
dim ids,iids
ids=split(Request.Form("seltodel"),",")
for each iids in ids
	if isnumeric(iids)=false or iids="" then
		Response.Redirect "admin.asp?user=" &ruser
		Response.End
	end if
next

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

cn.Execute Replace(Replace(sql_adminmpass,"{0}",Request.Form("seltodel")),"{1}",adminid),,1
cn.Close : set rs=nothing : set cn=nothing
%>
<!-- #include file="include/template/admin_traceback.inc" -->
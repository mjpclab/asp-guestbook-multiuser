<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_mpassaudit.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
end if
if Request.Form("seltodel")<>"" then
	dim ids
	ids=FilterSql(Request.Form("seltodel"))
	if Left(ids,1) = "," then
		ids=Mid(ids,2)
	end if

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)

	cn.Execute Replace(Replace(sql_adminmpass,"{0}",ids),"{1}",adminid),,1
	cn.Close : set rs=nothing : set cn=nothing
end if
%>
<!-- #include file="include/template/admin_traceback.inc" -->
<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_delfilter.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

if isnumeric(Request.Form("filterid")) and Request.Form("filterid")<>"" then
	 tfilterid=clng(Request.Form("filterid"))

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype

	cn.Execute Replace(Replace(sql_admindelfilter,"{0}",tfilterid),"{1}",adminid),,1
	
	cn.Close : set rs=nothing : set cn=nothing
end if

Response.Redirect "admin_filter.asp?user=" &ruser
%>
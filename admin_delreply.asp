<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_delreply.asp" -->
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
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp?user=" &ruser
	Response.End 
end if

set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

cn.BeginTrans
	cn.Execute Replace(Replace(sql_admindelreply_delete,"{0}",Request.QueryString("id")),"{1}",adminid),,129
	cn.Execute Replace(Replace(sql_admindelreply_update,"{0}",Request.QueryString("id")),"{1}",adminid),,129
cn.CommitTrans

cn.Close : set cn=nothing
%>
<!-- #include file="include/template/admin_traceback.inc" -->
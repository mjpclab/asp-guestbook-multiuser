<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_noreplyflag.asp" -->
<!-- #include file="include/sql/web_admin_searchdelreply.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "web_admin.asp"
	Response.End 
end if

set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

cn.BeginTrans
	cn.Execute sql_websearchdelreply_reply & Request.QueryString("id"),,129
	cn.Execute sql_websearchdelreply_unsetreply & Request.QueryString("id"),,129
cn.CommitTrans

cn.close
set cn=nothing
%>
<!-- #include file="include/template/web_admin_traceback.inc" -->
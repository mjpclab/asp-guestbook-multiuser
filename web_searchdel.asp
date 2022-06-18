<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_noreplyflag.asp" -->
<!-- #include file="include/sql/web_admin_searchdel.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="webconfig.asp" -->
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
	cn.Execute Replace(sql_webnoguestreply_flag,"{0}",Request.QueryString("id")),,129
	cn.Execute Replace(sql_websearchdel_reply,"{0}",Request.QueryString("id")),,129
	cn.Execute Replace(sql_websearchdel_main,"{0}",Request.QueryString("id")),,129
cn.CommitTrans

cn.close : set cn=nothing
%>
<!-- #include file="include/template/web_admin_traceback.inc" -->
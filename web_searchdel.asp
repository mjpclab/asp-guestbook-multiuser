<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="web_common.asp" -->

<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "web_admin.asp"
	Response.End 
end if

set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

cn.BeginTrans
	cn.Execute Replace(sql_global_webnoguestreply_flag,"{0}",Request.QueryString("id")),,1
	cn.Execute Replace(sql_websearchdel_reply,"{0}",Request.QueryString("id")),,1
	cn.Execute Replace(sql_websearchdel_main,"{0}",Request.QueryString("id")),,1
cn.CommitTrans

cn.close : set cn=nothing
%>
<!-- #include file="include/web_admin_traceback.inc" -->
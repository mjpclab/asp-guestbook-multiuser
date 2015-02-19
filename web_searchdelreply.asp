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
	cn.Execute sql_websearchdelreply_reply & Request.QueryString("id"),,1
	cn.Execute sql_websearchdelreply_unsetreply & Request.QueryString("id"),,1
cn.CommitTrans

cn.close
set cn=nothing


param_str=get_param_str()

Response.Redirect "web_search.asp" & param_str
%>
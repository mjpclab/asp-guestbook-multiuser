<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_noreplyflag.asp" -->
<!-- #include file="include/sql/web_admin_searchmdel.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
if Request.Form("seltodel")="" then
	Response.Redirect "web_search.asp"
	Response.End
end if
dim ids,iids
ids=split(Request.Form("seltodel"),",")
for each iids in ids
	if isnumeric(iids)=false or iids="" then
		Response.Redirect "web_admin.asp"
		Response.End
	end if
next

set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype


cn.BeginTrans
	cn.Execute Replace(sql_webnoguestreply_flag,"{0}",Request.Form("seltodel")),,1
	cn.Execute Replace(sql_websearchmdel_reply,"{0}",Request.Form("seltodel")),,1
	cn.Execute Replace(sql_websearchmdel_main,"{0}",Request.Form("seltodel")),,1
cn.CommitTrans

cn.close : set cn=nothing
%>
<!-- #include file="include/template/web_admin_traceback.inc" -->
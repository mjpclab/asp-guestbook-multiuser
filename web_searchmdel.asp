<!-- #include file="webconfig.asp" -->
<!-- #include file="inc_web_admin_stylesheet.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="web_common.asp" -->
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
	cn.Execute Replace(sql_global_webnoguestreply_flag,"{0}",Request.Form("seltodel")),,1
	cn.Execute Replace(sql_websearchmdel_reply,"{0}",Request.Form("seltodel")),,1
	cn.Execute Replace(sql_websearchmdel_main,"{0}",Request.Form("seltodel")),,1
cn.CommitTrans

cn.close : set cn=nothing
%>
<!-- #include file="include/web_admin_traceback.inc" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
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
checkuser cn,rs,true

cn.BeginTrans
	cn.Execute Replace(Replace(sql_global_noguestreply_flag,"{0}",Request.Form("seltodel")),"{1}",adminid),,1
	cn.Execute Replace(Replace(sql_adminmdel_reply,"{0}",Request.Form("seltodel")),"{1}",adminid),,1
	cn.Execute Replace(Replace(sql_adminmdel_main,"{0}",Request.Form("seltodel")),"{1}",adminid),,1
cn.CommitTrans

cn.Close : set rs=nothing : set cn=nothing
%>
<!-- #include file="admin_traceback.inc" -->
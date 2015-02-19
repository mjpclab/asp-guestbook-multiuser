<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if Request.Form("seltodel")="" then
	if isnumeric(Request.Form("tpage")) and Request.Form("tpage")<>"" then
		Response.Redirect "admin.asp?user=" &Request.Form("user")& "&page=" & Request.Form("tpage")
	else
		Response.Redirect "admin.asp?user=" &Request.Form("user")
	end if
end if
dim ids,iids
ids=split(Request.Form("seltodel"),",")
for each iids in ids
	if isnumeric(iids)=false or iids="" then
		Response.Redirect "admin.asp?user=" &Request.Form("user")
		Response.End
	end if
next

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,true

cn.Execute Replace(Replace(sql_adminmpass,"{0}",Request.Form("seltodel")),"{1}",Request.Form("user")),,1
cn.Close : set rs=nothing : set cn=nothing
%>
<!-- #include file="admin_traceback.inc" -->
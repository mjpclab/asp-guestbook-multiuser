<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1

if isnumeric(Request.Form("filterid")) and Request.Form("filterid")<>"" then
	tfilterid=clng(Request.Form("filterid"))

	set cn=server.CreateObject("ADODB.Connection")
	CreateConn cn,dbtype

	cn.Execute Replace(Replace(sql_admindelfilter,"{0}",tfilterid),"{1}",wm_id),,1
	
	cn.Close
	set cn=nothing
end if

Response.Redirect "web_filter.asp"
%>
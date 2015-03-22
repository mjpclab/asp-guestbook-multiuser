<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

if isnumeric(Request.Form("filterid")) and Request.Form("filterid")<>"" then
	 tfilterid=clng(Request.Form("filterid"))

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype
	checkuser cn,rs,true

	cn.Execute Replace(Replace(sql_admindelfilter,"{0}",tfilterid),"{1}",adminid),,1
	
	cn.Close : set rs=nothing : set cn=nothing
end if

Response.Redirect "admin_filter.asp?user=" &ruser
%>
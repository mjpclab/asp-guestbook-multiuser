<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp?user=" &ruser
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,false

cn.Execute Replace(Replace(Replace(sql_adminbring2top,"{0}",now()),"{1}",Request.QueryString("id")),"{2}",adminid),,1
cn.close : set rs=nothing : set cn=nothing
%>
<!-- #include file="admin_traceback.inc" -->
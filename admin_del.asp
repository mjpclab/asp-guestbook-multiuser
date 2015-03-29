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

cn.BeginTrans
	cn.Execute Replace(Replace(sql_global_noguestreply_flag,"{0}",Request.QueryString("id")),"{1}",adminid),,1
	cn.Execute Replace(Replace(sql_admindel_reply,"{0}",Request.QueryString("id")),"{1}",adminid),,1
	cn.Execute Replace(Replace(sql_admindel_main,"{0}",Request.QueryString("id")),"{1}",adminid),,1
cn.CommitTrans

cn.close : set rs=nothing : set cn=nothing
%>
<!-- #include file="admin_traceback.inc" -->
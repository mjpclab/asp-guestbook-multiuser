<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
	
CreateConn cn,dbtype
checkuser cn,rs,true

tnow=now()
cn.Execute Replace(Replace(sql_adminclearstats_startdate,"{0}",tnow),"{1}",adminid),,1
cn.Execute Replace(sql_adminclearstats_client,"{0}",adminid),,1

cn.close
set rs=nothing
set cn=nothing
		
Response.Redirect "admin_stats.asp?user=" &ruser
%>
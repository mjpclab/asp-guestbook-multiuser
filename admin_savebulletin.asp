<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusUserBanned then
	Response.Redirect "err.asp?user=" &ruser& "&number=100"
	Response.End
end if

set cn1=server.CreateObject("ADODB.Connection")
set rs1=server.CreateObject("ADODB.Recordset")
CreateConn cn1,dbtype
checkuser cn1,rs1,true

dim tlimit
tlimit=0
if Request.Form("html2")="1" then tlimit=tlimit+1
if Request.Form("ubb2")="1" then tlimit=tlimit+2
if Request.Form("newline2")="1" then tlimit=tlimit+4

dim tbul
tbul=Request.Form("abulletin")
tbul=replace(tbul,"'","''")
tbul=replace(tbul,"<%","< %")

cn1.Execute Replace(Replace(Replace(sql_adminsavebulletin,"{0}",clng(web_adminlimit and tlimit)),"{1}",tbul),"{2}",adminid),,1

cn1.close : set rs1=nothing : set cn1=nothing:
Response.Redirect "admin.asp?user=" &ruser
%>
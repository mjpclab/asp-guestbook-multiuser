<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_savebulletin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
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
CreateConn cn1,dbtype

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

cn1.close : set cn1=nothing:
Response.Redirect "admin.asp?user=" &ruser
%>
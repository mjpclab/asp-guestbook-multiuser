<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_savebulletin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
set cn1=server.CreateObject("ADODB.Connection")
Call CreateConn(cn1)

dim tlimit
tlimit=0
if Request.Form("html2")="1" then tlimit=tlimit+1
if Request.Form("ubb2")="1" then tlimit=tlimit+2
if Request.Form("newline2")="1" then tlimit=tlimit+4

if Request.Form("pub_at_face")="1" then tlimit=tlimit+16
if Request.Form("pub_at_function")="1" then tlimit=tlimit+32
if Request.Form("pub_at_index")="1" then tlimit=tlimit+64
if Request.Form("pub_at_search")="1" then tlimit=tlimit+128

dim tbul
tbul=Request.Form("abulletin")
tbul=replace(tbul,"'","''")
tbul=replace(tbul,"<%","< %")

cn1.Execute Replace(Replace(sql_websavebulletin,"{0}",tlimit),"{1}",tbul),,1

cn1.close
set cn1=nothing

Response.Redirect "web_admin.asp"
%>
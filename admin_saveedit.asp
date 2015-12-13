<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_saveedit.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
end if
if isnumeric(request.form("mainid"))=false or request.form("mainid")="" then
	Response.Redirect "admin.asp?user=" &ruser
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype


dim tlimit
tlimit=0
if Request.Form("html1")="1" then tlimit=tlimit+1
if Request.Form("ubb1")="1" then tlimit=tlimit+2
if Request.Form("newline1")="1" then tlimit=tlimit+4
	
rs.Open Replace(Replace(sql_adminsaveedit_open,"{0}",Request.Form("mainid")),"{1}",adminid),cn,0,3,1

if rs.EOF=false then		'ÁôÑÔ´æÔÚ
	rs.Fields("guestflag")=cbyte(rs.Fields("guestflag") and 248) + cbyte(tlimit and web_adminlimit)
	if Request.Form("etitle")<>"" then rs.Fields("title")=Server.HTMLEncode(Request.Form("etitle"))
	rs.Fields("article")=Request.Form("econtent")
	rs.Update

	%><!-- #include file="include/template/admin_traceback.inc" --><%
	rs.close : cn.close : set rs=nothing : set cn=nothing
else
	rs.close : cn.close : set rs=nothing : set cn=nothing
	Response.Redirect "admin.asp?user=" & ruser
end if
rs.close : cn.close : set rs=nothing : set cn=nothing
%>
<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<%
Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

Call CreateConn(cn)
checkuser cn,rs,false

if Request.QueryString("modeflag")="guest" then
	SetTimelessCookie InstanceName & "_DisplayMode_" & ruser ,Request.QueryString("mode")
elseif Request.QueryString("modeflag")="admin" then
	SetTimelessCookie InstanceName & "_AdminDisplayMode_" & ruser ,Request.QueryString("mode")
end if

if Request("rpage")<>"" then
	dim requests
	requests=HtmlDecode(GetRequests)
	if requests<>"" then requests="?" & mid(requests,2)

	Response.Redirect Request("rpage") & requests
end if
%>
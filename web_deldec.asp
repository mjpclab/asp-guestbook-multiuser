<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_deldec.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1

set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

cn.Execute Replace(sql_webdeldec,"{0}",FilterKeyword(Request.QueryString("user"))),,129
cn.close
set cn=nothing

if isnumeric(request.QueryString("page")) and request.QueryString("page")<>"" then
	Response.Redirect "web_searchdec.asp?adminname=" &server.URLEncode(request.QueryString("adminname"))& "&searchtxt=" & server.URLEncode(request.QueryString("searchtxt")) & "&page=" & request.QueryString("page")
else
	Response.Redirect "web_searchdec.asp?adminname=" &server.URLEncode(request.QueryString("adminname"))& "&searchtxt=" & server.URLEncode(request.QueryString("searchtxt"))
end if
%>
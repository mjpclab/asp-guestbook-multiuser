<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%
Response.Expires=-1

set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

cn.Execute Replace(sql_webdeldec,"{0}",replace(replace(ruser,"'",""),";","")),,1
cn.close
set cn=nothing

if isnumeric(request.QueryString("page")) and request.QueryString("page")<>"" then
	Response.Redirect "web_searchdec.asp?adminname=" &server.URLEncode(request.QueryString("adminname"))& "&searchtxt=" & server.URLEncode(request.QueryString("searchtxt")) & "&page=" & request.QueryString("page")
else
	Response.Redirect "web_searchdec.asp?adminname=" &server.URLEncode(request.QueryString("adminname"))& "&searchtxt=" & server.URLEncode(request.QueryString("searchtxt"))
end if
%>
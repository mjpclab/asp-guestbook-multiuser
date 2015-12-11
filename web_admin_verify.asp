<!-- #include file="include/sql/web_admin_verify.asp" -->
<%
Response.Expires=-1
Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

rs.Open sql_web_admin_verify,cn,,,1

if not rs.eof then
	if session.Contents(InstanceName & "_webpass")<>rs(0) then
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.Redirect "face.asp"
		Response.End
	else
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
	end if
else
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "face.asp"
	Response.End
end if
%>
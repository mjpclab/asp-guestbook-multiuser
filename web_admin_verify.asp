<!-- #include file="include/sql/web_admin_verify.asp" -->
<%
Response.Expires=-1
Dim cn,rs,refPass
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

rs.Open sql_web_admin_verify,cn,,,1
if Not rs.EOF then
	refPass=rs.Fields(0)
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing

if refPass="" Or Session(InstanceName & "_webpass")<>refPass then
	Response.Redirect "face.asp"
	Response.End
end if
%>
<!-- #include file="include/sql/web_admin_verify.asp" -->
<%
Sub VerifyFailed
	Response.Redirect "face.asp"
	Response.End
End Sub

Response.Expires=-1
Dim iadminpass
iadminpass=Session(InstanceName & "_webpass")
if iadminpass="" then
	Call VerifyFailed
else
	Dim cn,rs,refPass
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)

	rs.Open sql_web_admin_verify,cn,,,1
	if Not rs.EOF then
		refPass=rs.Fields(0)
	end if
	rs.Close : cn.Close : set rs=nothing : set cn=nothing

	if iadminpass<>refPass then
		Call VerifyFailed
	end if
end if
%>
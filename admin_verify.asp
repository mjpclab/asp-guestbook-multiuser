<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusLogin=false then
	Response.Redirect "web_err.asp?number=3"
	Response.End
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,false

rs.Open Replace(sql_adminverify,"{0}",adminid),cn,0,1,1

if not rs.eof then
	if session.Contents(InstanceName & "_adminpass_" & ruser)<>rs(0) then
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.Redirect "admin_login.asp?user=" &ruser
		Response.End
	else
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
	end if
else
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "admin_login.asp?user=" &ruser
	Response.End
end if
%>
<!-- #include file="web_error.asp" -->
<%
Function GetReferrer
	Dim QueryString
	QueryString=Request.QueryString
	if QueryString="" then
		GetReferrer = Request.ServerVariables("URL")
	else
		GetReferrer = Request.ServerVariables("URL") & "?" & Request.QueryString
	end if
End Function

Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif Not StatusLogin then
	Call WebErrorPage(3)
	Response.End
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
checkuser cn,rs,false

rs.Open Replace(sql_adminverify,"{0}",adminid),cn,0,1,1

if not rs.eof then
	if Session(InstanceName & "_adminpass_" & ruser)<>rs(0) then
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.Redirect "admin_login.asp?user=" &ruser& "&referrer=" & Server.UrlEncode(GetReferrer())
		Response.End
	else
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
	end if
else
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "admin_login.asp?user=" &ruser& "&referrer=" & Server.UrlEncode(GetReferrer())
	Response.End
end if
%>
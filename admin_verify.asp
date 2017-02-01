<!-- #include file="web_error.asp" -->
<!-- #include file="error.asp" -->
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

if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif Not StatusAccountEnabled then
	Call WebErrorPage(6)
	Response.End
elseif Not StatusAccountLoginEnabled then
	Call ErrorPage(101)
	Response.End
elseif Not StatusLogin then
	Call WebErrorPage(3)
	Response.End
end if

Sub VerifyFailed
	Response.Redirect "admin_login.asp?user=" &ruser& "&referrer=" & Server.UrlEncode(GetReferrer())
	Response.End
End Sub

Response.Expires=-1
Dim iadminpass
iadminpass=Session(InstanceName & "_adminpass_" & ruser)
if iadminpass="" then
	Call VerifyFailed
else
	Dim cn,rs,refPass
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	checkuser cn,rs,false

	rs.Open Replace(sql_adminverify,"{0}",adminid),cn,0,1,1
	if Not rs.EOF then
		refPass=rs.Fields(0)
	end if
	rs.Close : cn.Close : set rs=nothing : set cn=nothing

	if iadminpass<>refPass then
		Call VerifyFailed
	end if
end if
%>
<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="tips.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif Not StatusLogin then
	Call WebErrorPage(3)
	Response.End
end if

if StatusStatistics then call addstat("login")

Dim referrer
referrer=Request.Form("referrer")

if VcodeCount>0 then
	if VcodeCount>Len(Session(InstanceName & "_vcode")) then
		Session(InstanceName & "_vcode")=""
		if StatusStatistics then call addstat("loginfailed")
		Call TipsPage("验证码长度不足。","admin_login.asp?user=" &ruser& "&referrer=" & Server.UrlEncode(referrer))

		set rs=nothing
		set cn=nothing
		Response.End
	elseif Request.Form("ivcode")<>Session(InstanceName & "_vcode") or Session(InstanceName & "_vcode")="" then
		Session(InstanceName & "_vcode")=""
		if StatusStatistics then call addstat("loginfailed")
		Call TipsPage("验证码错误。","admin_login.asp?user=" &ruser& "&referrer=" & Server.UrlEncode(referrer))

		set rs=nothing
		set cn=nothing
		Response.End
	else
		Session(InstanceName & "_vcode")=""
	end if
end if

Sub LoginFailed
	if StatusStatistics then call addstat("loginfailed")
	Call TipsPage("密码不正确。","admin_login.asp?user=" &ruser& "&referrer=" & Server.UrlEncode(referrer))
End Sub

Dim iadminpass
iadminpass=Request.Form("iadminpass")
if iadminpass="" then
	Call LoginFailed
else
	iadminpass=md5(iadminpass,32)

	Dim cn,rs,refPass
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	rs.Open Replace(sql_adminverify,"{0}",adminid),cn,0,1,1
	if Not rs.EOF then
		refPass=rs.Fields(0)
	end if
	rs.Close : set rs=nothing

	if iadminpass=refPass then
		cn.Execute Replace(Replace(sql_updatelastlogin,"{0}",ServerTimeToUTC(now())),"{1}",adminid),,129
		cn.Close : set cn=nothing

		Session(InstanceName & "_adminpass_"& ruser)=iadminpass
		session.Timeout=AdminTimeOut
		if referrer<>"" then
			Response.Redirect referrer
		else
			Response.Redirect "admin.asp?user=" &ruser
		end if
	else
		cn.Close : set cn=nothing
		Call LoginFailed
	end if
end if
%>
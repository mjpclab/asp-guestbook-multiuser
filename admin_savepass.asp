<!-- #include file="loadconfig.asp" -->
<link rel="stylesheet" type="text/css" href="style.css"/>
<!-- #include file="style.asp" -->
<!-- #include file="md5.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,true

if request.Form("ioldpass")="" then
	Call MessagePage("密码不能为空，请重新输入。","admin_chpass.asp?user=" &Request.Form("user"))
elseif request.Form("inewpass1")<> request.Form("inewpass2") then
	Call MessagePage("新密码不一致，请检查。","admin_chpass.asp?user=" &Request.Form("user"))
elseif request.Form("inewpass1")="" then
	Call MessagePage("新密码不能为空，请重新输入。","admin_chpass.asp?user=" &Request.Form("user"))
else
	rs.Open Replace(sql_adminsavepass_select,"{0}",Request.Form("user")),cn,0,1,1

	if session.Contents(InstanceName & "_adminpass_" & Request.Form("user"))=rs(0) then
		if md5(request.Form("ioldpass"),32)=session.Contents(InstanceName & "_adminpass_" & Request.Form("user")) then
			pwd=md5(request.Form("inewpass1"),32)
					
			cn.Execute Replace(Replace(sql_adminsavepass_update,"{0}",pwd),"{1}",Request.Form("user")),,1
			session.Contents(InstanceName & "_adminpass_" & Request.Form("user"))=pwd
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			Response.Redirect "admin.asp?user=" &Request.Form("user")
		else
			Call MessagePage("原密码错误，请检查。","admin_chpass.asp?user=" &Request.Form("user"))
		end if
	else
		Call MessagePage("原密码验证失败，请重新登录后再试。","admin_chpass.asp?user=" &Request.Form("user"))
	end if
	
	rs.Close
end if
cn.Close : set rs=nothing : set cn=nothing
%>
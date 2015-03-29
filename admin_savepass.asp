<!-- #include file="loadconfig.asp" -->
<link rel="stylesheet" type="text/css" href="css/style.css"/>
<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
<!-- #include file="css/style.asp" -->
<!-- #include file="css/adminstyle.asp" -->
<!-- #include file="md5.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,true

if request.Form("ioldpass")="" then
	Call MessagePage("密码不能为空，请重新输入。","admin_chpass.asp?user=" &ruser)
elseif request.Form("inewpass1")<> request.Form("inewpass2") then
	Call MessagePage("新密码不一致，请检查。","admin_chpass.asp?user=" &ruser)
elseif request.Form("inewpass1")="" then
	Call MessagePage("新密码不能为空，请重新输入。","admin_chpass.asp?user=" &ruser)
else
	rs.Open Replace(sql_adminsavepass_select,"{0}",adminid),cn,0,1,1

	if session.Contents(InstanceName & "_adminpass_" & ruser)=rs(0) then
		if md5(request.Form("ioldpass"),32)=session.Contents(InstanceName & "_adminpass_" & ruser) then
			pwd=md5(request.Form("inewpass1"),32)
					
			cn.Execute Replace(Replace(sql_adminsavepass_update,"{0}",pwd),"{1}",adminid),,1
			session.Contents(InstanceName & "_adminpass_" & ruser)=pwd
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			Response.Redirect "admin.asp?user=" &ruser
		else
			Call MessagePage("原密码错误，请检查。","admin_chpass.asp?user=" &ruser)
		end if
	else
		Call MessagePage("原密码验证失败，请重新登录后再试。","admin_chpass.asp?user=" &ruser)
	end if
	
	rs.Close
end if
cn.Close : set rs=nothing : set cn=nothing
%>
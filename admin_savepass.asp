<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_savepass.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

if request.Form("ioldpass")="" then
	Call TipsPage("密码不能为空，请重新输入。","admin_chpass.asp?user=" &ruser)
elseif request.Form("inewpass1")<> request.Form("inewpass2") then
	Call TipsPage("新密码不一致，请检查。","admin_chpass.asp?user=" &ruser)
elseif request.Form("inewpass1")="" then
	Call TipsPage("新密码不能为空，请重新输入。","admin_chpass.asp?user=" &ruser)
else
	rs.Open Replace(sql_adminsavepass_select,"{0}",adminid),cn,0,1,1

	if Session(InstanceName & "_adminpass_" & ruser)=rs(0) then
		if md5(request.Form("ioldpass"),32)=Session(InstanceName & "_adminpass_" & ruser) then
			pwd=md5(request.Form("inewpass1"),32)
					
			cn.Execute Replace(Replace(sql_adminsavepass_update,"{0}",pwd),"{1}",adminid),,1
			Session(InstanceName & "_adminpass_" & ruser)=pwd
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			Response.Redirect "admin.asp?user=" &ruser
		else
			Call TipsPage("原密码错误，请检查。","admin_chpass.asp?user=" &ruser)
		end if
	else
		Call TipsPage("原密码验证失败，请重新登录后再试。","admin_chpass.asp?user=" &ruser)
	end if
	
	rs.Close
end if
cn.Close : set rs=nothing : set cn=nothing
%>
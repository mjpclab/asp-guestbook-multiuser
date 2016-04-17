<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_saveqa.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/string.asp" -->
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

if IsEmpty(Request.Form) then
	Response.Redirect "admin_chpass.asp?user=" &ruser
elseif request.Form("ioldpass")="" then
	Call TipsPage("密码不能为空，请重新输入。","admin_chpass.asp?user=" &ruser)
elseif request.Form("question")="" then
	Call TipsPage("找回密码问题不能为空，请重新输入。","admin_chpass.asp?user=" &ruser)
elseif request.Form("key")="" then
	Call TipsPage("找回密码答案不能为空，请重新输入。","admin_chpass.asp?user=" &ruser)
else
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	rs.Open Replace(sql_adminsaveqa_select,"{0}",adminid),cn,,,1

	if Session(InstanceName & "_adminpass_" & ruser)=rs(0) then
		if md5(request.Form("ioldpass"),32)=Session(InstanceName & "_adminpass_" & ruser) then
			cn.Execute Replace(Replace(Replace(sql_adminsaveqa_update,"{0}",replace(Request.Form("question"),"'","''")),"{1}",md5(Request.Form("key"),32)),"{2}",adminid),,1
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			Response.Redirect "admin.asp?user=" &ruser
		else
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			Call TipsPage("密码错误，请检查。","admin_chpass.asp?user=" &ruser)
		end if
	else
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Call TipsPage("原密码验证失败，请重新登录后再试。","admin_chpass.asp?user=" &ruser)
	end if
end if
%>
<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_savepass.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1

if IsEmpty(Request.Form) then
	Response.Redirect "web_chpass.asp"
elseif request.Form("ioldpass")="" then
	Call TipsPage("原密码不能为空，请重新输入。","web_chpass.asp")
elseif request.Form("inewpass1")<> request.Form("inewpass2") then
	Call TipsPage("新密码不一致，请检查。","web_chpass.asp")
elseif request.Form("inewpass1")="" then
	Call TipsPage("密码不能为空，请重新输入。","web_chpass.asp")
else
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	rs.Open sql_websavepass_select,cn,,,1

	if md5(request.Form("ioldpass"),32)=rs(0) then
		pwd=md5(request.Form("inewpass1"),32)

		cn.Execute Replace(sql_websavepass_update,"{0}",pwd),,129
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Session(InstanceName & "_webpass")=pwd
		Response.Redirect "web_admin.asp"
	else
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Call TipsPage("原密码错误，请检查。","web_chpass.asp")
	end if
end if
%>
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
	Call TipsPage("���벻��Ϊ�գ����������롣","admin_chpass.asp?user=" &ruser)
elseif request.Form("inewpass1")<> request.Form("inewpass2") then
	Call TipsPage("�����벻һ�£����顣","admin_chpass.asp?user=" &ruser)
elseif request.Form("inewpass1")="" then
	Call TipsPage("�����벻��Ϊ�գ����������롣","admin_chpass.asp?user=" &ruser)
else
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	rs.Open Replace(sql_adminsavepass_select,"{0}",adminid),cn,0,1,1

	if Session(InstanceName & "_adminpass_" & ruser)=rs(0) then
		if md5(request.Form("ioldpass"),32)=Session(InstanceName & "_adminpass_" & ruser) then
			pwd=md5(request.Form("inewpass1"),32)

			cn.Execute Replace(Replace(sql_adminsavepass_update,"{0}",pwd),"{1}",adminid),,129
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			Session(InstanceName & "_adminpass_" & ruser)=pwd
			Response.Redirect "admin.asp?user=" &ruser
		else
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			Call TipsPage("ԭ����������顣","admin_chpass.asp?user=" &ruser)
		end if
	else
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Call TipsPage("ԭ������֤ʧ�ܣ������µ�¼�����ԡ�","admin_chpass.asp?user=" &ruser)
	end if
end if
%>
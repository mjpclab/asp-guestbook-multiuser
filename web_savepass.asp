<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_savepass.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

if request.Form("ioldpass")="" then
	Call TipsPage("ԭ���벻��Ϊ�գ����������롣","web_chpass.asp")
elseif request.Form("inewpass1")<> request.Form("inewpass2") then
	Call TipsPage("�����벻һ�£����顣","web_chpass.asp")
elseif request.Form("inewpass1")="" then
	Call TipsPage("���벻��Ϊ�գ����������롣","web_chpass.asp")
else
	rs.Open sql_websavepass_select,cn,,,1

	if md5(request.Form("ioldpass"),32)=rs(0) then
		pwd=md5(request.Form("inewpass1"),32)
				
		cn.Execute Replace(sql_websavepass_update,"{0}",pwd),,1
		Session(InstanceName & "_webpass")=pwd
		
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.Redirect "web_admin.asp"
	else
		Call TipsPage("ԭ����������顣","web_chpass.asp")
	end if

	rs.Close
end if
cn.Close : set rs=nothing : set cn=nothing
%>
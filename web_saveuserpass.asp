<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/web_admin_saveuserpass.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

checkuser cn,rs,true

if request.Form("inewpass1")<> request.Form("inewpass2") then
	cn.Close : set rs=nothing : set cn=nothing
	Call TipsPage("�����벻һ�£����顣","web_chuserpass.asp?user=" & ruser)
elseif request.Form("inewpass1")="" then
	cn.Close : set rs=nothing : set cn=nothing
	Call TipsPage("���벻��Ϊ�գ����������롣","web_chuserpass.asp?user=" & ruser)
else
	cn.Execute Replace(Replace(sql_websaveuserpass,"{0}",md5(Request.Form("inewpass1"),32)),"{1}",ruser),,1
	cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "web_userinfo.asp?user=" & ruser
end if
%>
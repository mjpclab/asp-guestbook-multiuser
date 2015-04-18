<!-- #include file="webconfig.asp" -->
<!-- #include file="inc_web_admin_stylesheet.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="md5.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

checkuser cn,rs,true

if request.Form("inewpass1")<> request.Form("inewpass2") then
	Call MessagePage("新密码不一致，请检查。","web_chuserpass.asp?user=" & ruser)
elseif request.Form("inewpass1")="" then
	Call MessagePage("密码不能为空，请重新输入。","web_chuserpass.asp?user=" & ruser)
else
	cn.Execute Replace(Replace(sql_websaveuserpass,"{0}",md5(Request.Form("inewpass1"),32)),"{1}",ruser),,1
	cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "web_userinfo.asp?user=" & ruser
end if
cn.Close : set rs=nothing : set cn=nothing
%>
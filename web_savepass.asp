<!-- #include file="webconfig.asp" -->
<link rel="stylesheet" type="text/css" href="style.css"/>
<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
<link rel="stylesheet" type="text/css" href="web_adminstyle.css"/>
<!-- #include file="style.asp" -->
<!-- #include file="adminstyle.asp" -->
<!-- #include file="web_adminstyle.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="md5.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

if request.Form("ioldpass")="" then
	Call MessagePage("ԭ���벻��Ϊ�գ����������롣","web_chpass.asp")
elseif request.Form("inewpass1")<> request.Form("inewpass2") then
	Call MessagePage("�����벻һ�£����顣","web_chpass.asp")
elseif request.Form("inewpass1")="" then
	Call MessagePage("���벻��Ϊ�գ����������롣","web_chpass.asp")
else
	rs.Open sql_websavepass_select,cn,,,1

	if md5(request.Form("ioldpass"),32)=rs(0) then
		pwd=md5(request.Form("inewpass1"),32)
				
		cn.Execute Replace(sql_websavepass_update,"{0}",pwd),,1
		session.Contents(InstanceName & "_webpass")=pwd
		
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.Redirect "web_admin.asp"
	else
		Call MessagePage("ԭ����������顣","web_chpass.asp")
	end if

	rs.Close
end if
cn.Close : set rs=nothing : set cn=nothing
%>
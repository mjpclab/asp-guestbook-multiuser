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
	Call MessagePage("���벻��Ϊ�գ����������롣","admin_chpass.asp?user=" &ruser)
elseif request.Form("question")="" then
	Call MessagePage("�һ��������ⲻ��Ϊ�գ����������롣","admin_chpass.asp?user=" &ruser)
elseif request.Form("key")="" then
	Call MessagePage("�һ�����𰸲���Ϊ�գ����������롣","admin_chpass.asp?user=" &ruser)
else
	rs.Open Replace(sql_adminsaveqa_select,"{0}",adminid),cn,,,1
	if session.Contents(InstanceName & "_adminpass_" & ruser)=rs(0) then
		if md5(request.Form("ioldpass"),32)=session.Contents(InstanceName & "_adminpass_" & ruser) then
			
			cn.Execute Replace(Replace(Replace(sql_adminsaveqa_update,"{0}",replace(Request.Form("question"),"'","''")),"{1}",md5(Request.Form("key"),32)),"{2}",adminid),,1
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			Response.Redirect "admin.asp?user=" &ruser
		else
			Call MessagePage("����������顣","admin_chpass.asp?user=" &ruser)
		end if
	else
		Call MessagePage("ԭ������֤ʧ�ܣ������µ�¼�����ԡ�","admin_chpass.asp?user=" &ruser)
	end if
	
	rs.Close
end if

cn.Close : set rs=nothing : set cn=nothing
%>
<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

dim re
set re=new RegExp
re.Pattern="^\w+$"

if Request.QueryString("user")="" then
	Response.Write "�û�������Ϊ�ա�"
elseif re.Test(Request.QueryString("user"))=false then
	Response.Write "�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�"
else
	dim cn,rs
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype

	'=============================��������֤
	rs.Open Replace(sql_checkuser,"{0}",Request.QueryString("user")),cn,,,1
	if not rs.EOF then
		Response.Write "�û����� " & Request.QueryString("user") & " �Ѵ��ڣ��뻻�������û�����"
	else
		Response.Write "�û����� " & Request.QueryString("user") & " ����ʹ�á�"
	end if
end if
%>
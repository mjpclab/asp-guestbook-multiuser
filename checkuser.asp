<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

dim re
set re=new RegExp
re.Pattern="^\w+$"

ruser=Request.QueryString("user")
if ruser="" then
	Response.Write "�û�������Ϊ�ա�"
elseif re.Test(ruser)=false then
	Response.Write "�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�"
else
	dim cn,rs
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype

	'=============================��������֤
	rs.Open Replace(sql_checkuser,"{0}",ruser),cn,,,1
	if not rs.EOF then
		Response.Write "�û����� " & ruser & " �Ѵ��ڣ��뻻�������û�����"
	else
		Response.Write "�û����� " & ruser & " ����ʹ�á�"
	end if
end if
%>
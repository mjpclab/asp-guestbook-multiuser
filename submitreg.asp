<!-- #include file="webconfig.asp" -->
<!-- #include file="common2.asp" -->
<!-- #include file="md5.asp" -->
<%

Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusReg=false then
	Response.Redirect "web_err.asp?number=1"
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	Call MessagePage("��֤�����","reg.asp")
	Response.End
else
	session("vcode")=""
end if
'===============================��ʽ��֤
dim re
set re=new RegExp
re.Pattern="^\w+$"

if Request.Form("user")="" then
	Call MessagePage("�û�������Ϊ�ա�","reg.asp")
	Response.End
elseif re.Test(Request.Form("user"))=false then
	Call MessagePage("�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�","reg.asp")
	Response.End
elseif len(Request.Form("user"))>32 then
	Call MessagePage("�û������Ȳ��ܳ���32�֡�","reg.asp")
	Response.End
elseif Request.Form("pass1")="" then
	Call MessagePage("���벻��Ϊ�ա�","reg.asp")
	Response.End
elseif Request.Form("pass2")="" then
	Call MessagePage("ȷ�����벻��Ϊ�ա�","reg.asp")
	Response.End
elseif Request.Form("pass1")<>Request.Form("pass2") then
	Call MessagePage("���벻һ�£����顣","reg.asp")
	Response.End
elseif len(Request.Form("pass1"))>32 then
	Call MessagePage("���볤�Ȳ��ܳ���32�֡�","reg.asp")
	Response.End
elseif Request.Form("question")="" then
	Call MessagePage("�һ��������ⲻ��Ϊ�ա�","reg.asp")
	Response.End
elseif Request.Form("key")="" then
	Call MessagePage("�һ�����𰸲���Ϊ�ա�","reg.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

'=============================��������֤
rs.Open Replace(sql_submitreg_checkuser,"{0}",Request.Form("user")),cn,,,1
if not rs.EOF then
	Call MessagePage("�û����Ѵ��ڣ����������롣","reg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close

dim tnow
cn.BeginTrans
tnow=now()
cn.Execute Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_submitreg_init1,"{0}",Request.Form("user")),"{1}",md5(Request.Form("pass1"),32)),"{2}",FilterQuote(server.HTMLEncode(Request.Form("nick")))),"{3}",0),"{4}",tnow),"{5}",tnow),"{6}",FilterQuote(Request.Form("question"))),"{7}",md5(Request.Form("key"),32)),"{8}",""),,1
cn.Execute Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_submitreg_init2,"{0}",Request.Form("user")),"{1}",143),"{2}",6),"{3}",6),"{4}",20),"{5}",111),"{6}",1270),"{7}","����"),,1
cn.Execute Replace(sql_submitreg_init3,"{0}",Request.Form("user")),,1
cn.CommitTrans

set rs=nothing

dim gbookaddr
gbookaddr=geturlpath & "index.asp?user=" & Request.Form("user")
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> �����ɹ�</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<!-- #include file="css/style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">��ҳ</a> &gt;&gt; <a href="reg.asp" style="color:<%=TitleColor%>">�������Ա�</a> &gt;&gt; �����ɹ�</div>
</div>

<%
dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="func_web.inc" -->

<p class="centertext">�����ɹ���������������Ա���ҳ��ַ��<br/><%=gbookaddr%><br/>&gt;<a href="<%=gbookaddr%>">ת����ҳ��</a></p>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
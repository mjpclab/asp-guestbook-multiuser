<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/sql/web_submitreg.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="tips.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif StatusReg=false then
	Call WebErrorPage(1)
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>Session("vcode") or Session("vcode")="") then
	Session("vcode")=""
	Call TipsPage("��֤�����","reg.asp")
	Response.End
else
	Session("vcode")=""
end if
'===============================��ʽ��֤
dim re
set re=new RegExp
re.Pattern="^\w+$"

ruser=Request.Form("user")
if ruser="" then
	Call TipsPage("�û�������Ϊ�ա�","reg.asp")
	Response.End
elseif re.Test(ruser)=false then
	Call TipsPage("�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�","reg.asp")
	Response.End
elseif len(ruser)>32 then
	Call TipsPage("�û������Ȳ��ܳ���32�֡�","reg.asp")
	Response.End
elseif Request.Form("pass1")="" then
	Call TipsPage("���벻��Ϊ�ա�","reg.asp")
	Response.End
elseif Request.Form("pass2")="" then
	Call TipsPage("ȷ�����벻��Ϊ�ա�","reg.asp")
	Response.End
elseif Request.Form("pass1")<>Request.Form("pass2") then
	Call TipsPage("���벻һ�£����顣","reg.asp")
	Response.End
elseif len(Request.Form("pass1"))>32 then
	Call TipsPage("���볤�Ȳ��ܳ���32�֡�","reg.asp")
	Response.End
elseif Request.Form("question")="" then
	Call TipsPage("�һ��������ⲻ��Ϊ�ա�","reg.asp")
	Response.End
elseif Request.Form("key")="" then
	Call TipsPage("�һ�����𰸲���Ϊ�ա�","reg.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

'=============================��������֤
rs.Open Replace(sql_submitreg_checkuser,"{0}",ruser),cn,,,1
if not rs.EOF then
	Call TipsPage("�û����Ѵ��ڣ����������롣","reg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close

dim tnow
cn.BeginTrans
tnow=now()
cn.Execute Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_submitreg_init1,"{0}",ruser),"{1}",md5(Request.Form("pass1"),32)),"{2}",FilterQuote(server.HTMLEncode(Request.Form("nick")))),"{3}",0),"{4}",tnow),"{5}",tnow),"{6}",FilterQuote(Request.Form("question"))),"{7}",md5(Request.Form("key"),32)),"{8}",""),,1
cn.Execute Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_submitreg_init2,"{0}",ruser),"{1}",143),"{2}",6),"{3}",6),"{4}",20),"{5}",34),"{6}",132),"{7}",132),"{8}",107),"{9}",1270),"{10}",CssFontFamily),"{11}",CssFontSize),"{12}",CssLineHeight),"{13}",styleid),,1
cn.Execute Replace(sql_submitreg_init3,"{0}",ruser),,1
cn.CommitTrans

set rs=nothing

dim gbookaddr
gbookaddr=geturlpath & "index.asp?user=" & ruser
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> �����ɹ�</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">��ҳ</a> &gt;&gt; <a href="reg.asp" style="color:<%=TitleColor%>">�������Ա�</a> &gt;&gt; ����ɹ�</div>
</div>
<div id="mainborder" class="mainborder">
<%
dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="include/template/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/template/web_guest_func.inc" -->

<p class="centertext">�����ɹ���������������Ա���ҳ��ַ��<br/><%=gbookaddr%><br/>&gt;<a href="<%=gbookaddr%>">ת����ҳ��</a></p>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
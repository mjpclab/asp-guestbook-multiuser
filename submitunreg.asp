<!-- #include file="webconfig.asp" -->
<link rel="stylesheet" type="text/css" href="style.css"/>
<!-- #include file="style.asp" -->
<!-- #include file="md5.asp" -->
<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusUnreg=false then
	Response.Redirect "web_err.asp?number=5"
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	Call MessagePage("��֤�����","unreg.asp")
	Response.End
else
	session("vcode")=""
end if

'===============================��ʽ��֤
dim re
set re=new RegExp
re.Pattern="^\w+$"

if Request.Form("user")="" then
	Call MessagePage("�û�������Ϊ�ա�","unreg.asp")
	Response.End
elseif re.Test(Request.Form("user"))=false then
	Call MessagePage("�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�","unreg.asp")
	Response.End
elseif len(Request.Form("user"))>32 then
	Call MessagePage("�û������Ȳ��ܳ���32�֡�","unreg.asp")
	Response.End
elseif Request.Form("pass1")="" then
	Call MessagePage("���벻��Ϊ�ա�","unreg.asp")
	Response.End
elseif len(Request.Form("pass1"))>32 then
	Call MessagePage("���볤�Ȳ��ܳ���32�֡�","unreg.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

'=============================��������֤
rs.Open Replace(Replace(sql_submitunreg_checkuser,"{0}",Request.Form("user")),"{1}",wm_name),cn,,,1
if rs.EOF then
	Call MessagePage("�û��������ڡ�","unreg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
elseif rs("adminpass")<>md5(Request.Form("pass1"),32) then
	Call MessagePage("���벻��ȷ��","unreg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : set rs=nothing

cn.BeginTrans
	cn.Execute Replace(sql_submitunreg_delete_filterconfig,"{0}",Request.Form("user")),,1
	cn.Execute Replace(sql_submitunreg_delete_ipconfig,"{0}",Request.Form("user")),,1
	cn.Execute Replace(sql_submitunreg_delete_floodconfig,"{0}",Request.Form("user")),,1
	cn.Execute Replace(sql_submitunreg_delete_stats,"{0}",Request.Form("user")),,1
	cn.Execute Replace(sql_submitunreg_delete_stats_clientinfo,"{0}",Request.Form("user")),,1
	cn.Execute Replace(sql_submitunreg_delete_reply,"{0}",Request.Form("user")),,1
	cn.Execute Replace(sql_submitunreg_delete_main,"{0}",Request.Form("user")),,1
	cn.Execute Replace(sql_submitunreg_delete_config,"{0}",Request.Form("user")),,1
	cn.Execute Replace(sql_submitunreg_delete_supervisor,"{0}",Request.Form("user")),,1
cn.CommitTrans

cn.Close : set cn=nothing

Call MessagePage("�ʺ�ɾ����ɡ�","unreg.asp")
Response.Redirect "face.asp"
%>
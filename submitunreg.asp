<!-- #include file="webconfig.asp" -->
<!-- #include file="inc_stylesheet.asp" -->
<!-- #include file="include/md5.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP then
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

ruser=Request.Form("user")
if ruser="" then
	Call MessagePage("�û�������Ϊ�ա�","unreg.asp")
	Response.End
elseif re.Test(ruser)=false then
	Call MessagePage("�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�","unreg.asp")
	Response.End
elseif len(ruser)>32 then
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
Dim del_adminid
rs.Open Replace(Replace(sql_submitunreg_checkuser,"{0}",ruser),"{1}",wm_name),cn,,,1
if rs.EOF then
	Call MessagePage("�û��������ڡ�","unreg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
elseif rs("adminpass")<>md5(Request.Form("pass1"),32) then
	Call MessagePage("���벻��ȷ��","unreg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
del_adminid=rs.Fields("adminid")
rs.Close : set rs=nothing

cn.BeginTrans
	cn.Execute Replace(sql_submitunreg_delete_filterconfig,"{0}",del_adminid),,1
	cn.Execute Replace(sql_submitunreg_delete_ipv4config,"{0}",del_adminid),,1
	cn.Execute Replace(sql_submitunreg_delete_ipv6config,"{0}",del_adminid),,1
	cn.Execute Replace(sql_submitunreg_delete_floodconfig,"{0}",del_adminid),,1
	cn.Execute Replace(sql_submitunreg_delete_stats,"{0}",del_adminid),,1
	cn.Execute Replace(sql_submitunreg_delete_stats_clientinfo,"{0}",del_adminid),,1
	cn.Execute Replace(sql_submitunreg_delete_reply,"{0}",del_adminid),,1
	cn.Execute Replace(sql_submitunreg_delete_main,"{0}",del_adminid),,1
	cn.Execute Replace(sql_submitunreg_delete_config,"{0}",del_adminid),,1
	cn.Execute Replace(sql_submitunreg_delete_supervisor,"{0}",del_adminid),,1
cn.CommitTrans

cn.Close : set cn=nothing

Call MessagePage("�ʺ�ɾ����ɡ�","unreg.asp")
Response.Redirect "face.asp"
%>
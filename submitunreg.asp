<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/web_submitunreg.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="tips.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif Not StatusUnreg then
	Call WebErrorPage(5)
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>Session(InstanceName & "_vcode") or Session(InstanceName & "_vcode")="") then
	Session(InstanceName & "_vcode")=""
	Call TipsPage("验证码错误。","unreg.asp")
	Response.End
else
	Session(InstanceName & "_vcode")=""
end if

'===============================合式验证
dim re
set re=new RegExp
re.Pattern="^\w+$"

ruser=Request.Form("user")
if ruser="" then
	Call TipsPage("用户名不能为空。","unreg.asp")
	Response.End
elseif re.Test(ruser)=false then
	Call TipsPage("用户名中只能包含英文字母、数字和下划线。","unreg.asp")
	Response.End
elseif len(ruser)>32 then
	Call TipsPage("用户名长度不能超过32字。","unreg.asp")
	Response.End
elseif Request.Form("pass1")="" then
	Call TipsPage("密码不能为空。","unreg.asp")
	Response.End
elseif len(Request.Form("pass1"))>32 then
	Call TipsPage("密码长度不能超过32字。","unreg.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

'=============================存在性验证
Dim del_adminid
rs.Open Replace(Replace(sql_submitunreg_checkuser,"{0}",ruser),"{1}",wm_name),cn,,,1
if rs.EOF then
	Call TipsPage("用户名不存在。","unreg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
elseif rs("adminpass")<>md5(Request.Form("pass1"),32) then
	Call TipsPage("密码不正确。","unreg.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
del_adminid=rs.Fields("adminid")
rs.Close : set rs=nothing

cn.BeginTrans
	cn.Execute Replace(sql_submitunreg_delete_filterconfig,"{0}",del_adminid),,129
	cn.Execute Replace(sql_submitunreg_delete_ipv4config,"{0}",del_adminid),,129
	cn.Execute Replace(sql_submitunreg_delete_ipv6config,"{0}",del_adminid),,129
	cn.Execute Replace(sql_submitunreg_delete_floodconfig,"{0}",del_adminid),,129
	cn.Execute Replace(sql_submitunreg_delete_stats,"{0}",del_adminid),,129
	cn.Execute Replace(sql_submitunreg_delete_stats_clientinfo,"{0}",del_adminid),,129
	cn.Execute Replace(sql_submitunreg_delete_reply,"{0}",del_adminid),,129
	cn.Execute Replace(sql_submitunreg_delete_main,"{0}",del_adminid),,129
	cn.Execute Replace(sql_submitunreg_delete_config,"{0}",del_adminid),,129
	cn.Execute Replace(sql_submitunreg_delete_supervisor,"{0}",del_adminid),,129
cn.CommitTrans

cn.Close : set cn=nothing

Call TipsPage("帐号删除完成。","face.asp")
%>
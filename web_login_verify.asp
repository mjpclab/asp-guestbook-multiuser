<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/sql/web_admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="tips.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if VcodeCount>0 and (Request.Form("ivcode")<>Session("vcode") or Session("vcode")="") then
	Session("vcode")=""
	Call TipsPage("��֤�����","web_login.asp")
	Response.End
else
	Session("vcode")=""
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

Call CreateConn(cn)
rs.Open sql_web_admin_verify,cn,0,1,1

Session(InstanceName & "_webpass")=md5(request("iadminpass"),32)
if Session(InstanceName & "_webpass")=rs(0) then
	session.Timeout=cint(AdminTimeOut)
	
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "web_admin.asp"
else
	Call TipsPage("���벻��ȷ��","web_login.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
end if
%>
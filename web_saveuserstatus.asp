<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/web_admin_saveuserstatus.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

checkuser cn,rs,true

if request.Form("statusDisabled")="0" then
	cn.Execute Replace(sql_web_enableuser,"{0}",adminid),,129
else
	cn.Execute Replace(sql_web_disableuser,"{0}",adminid),,129
end if

if request.Form("statusLoginDisabled")="0" then
	cn.Execute Replace(sql_web_enableuserlogin,"{0}",adminid),,129
else
	cn.Execute Replace(sql_web_disableuserlogin,"{0}",adminid),,129
end if

if request.Form("statusLeavewordDisabled")="0" then
	cn.Execute Replace(sql_web_enableuserleaveword,"{0}",adminid),,129
else
	cn.Execute Replace(sql_web_disableuserleaveword,"{0}",adminid),,129
end if

cn.Close : set rs=nothing : set cn=nothing
Response.Redirect "web_userinfo.asp?user=" & ruser
%>
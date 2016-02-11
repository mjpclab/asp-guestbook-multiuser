<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="loadconfig.asp" -->
<%
Session.Contents.Remove(InstanceName & "_adminpass_" & ruser)
Response.Redirect "admin_login.asp?user=" &ruser
%>
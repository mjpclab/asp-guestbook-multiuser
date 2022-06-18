<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
SetTimelessCookie InstanceName & "_WebDisplayMode_", Request.QueryString("mode")

param_str=get_param_str()
Response.Redirect "web_search.asp" & param_str
%>
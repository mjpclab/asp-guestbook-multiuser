<!-- #include file="loadconfig.asp" -->
<%
Session.Contents(InstanceName & "_adminpass_" & Request("user"))=empty
Response.Redirect "index.asp?user=" &Request.QueryString("user")
%>
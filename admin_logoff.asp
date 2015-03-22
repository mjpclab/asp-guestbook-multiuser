<!-- #include file="loadconfig.asp" -->
<%
Session.Contents(InstanceName & "_adminpass_" & ruser)=empty
Response.Redirect "index.asp?user=" &ruser
%>
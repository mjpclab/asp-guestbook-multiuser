<!-- #include file="webconfig.asp" -->
<%
Session.Contents(InstanceName & "_webpass")=empty
Response.Redirect "face.asp"
%>
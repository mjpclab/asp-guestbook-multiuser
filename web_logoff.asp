<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<%
Session.Contents(InstanceName & "_webpass")=empty
Response.Redirect "face.asp"
%>
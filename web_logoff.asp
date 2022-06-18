<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<%
Session.Contents.Remove(InstanceName & "_webpass")
Response.Redirect "face.asp"
%>
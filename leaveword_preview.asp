<!-- #include file="loadconfig.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

dim PreviewContent
PreviewContent=Request.Form("icontent")
convertstr PreviewContent,guestlimit and web_guestlimit,3

Response.Write PreviewContent
%>
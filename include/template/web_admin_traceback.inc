<%
Dim CurrentDisplayMode
CurrentDisplayMode=WebDisplayMode()
if CurrentDisplayMode="book" then
	Response.Redirect "web_search.asp" & get_param_str()
elseif CurrentDisplayMode="forum" then
	Response.Redirect "web_showword.asp?id=" & Request("rootid") & "&" & mid(get_param_str(),2)
end if
%>
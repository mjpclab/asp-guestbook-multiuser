<%
if WebDisplayMode()="book" then
	Response.Redirect "web_search.asp" & get_param_str()
elseif WebDisplayMode()="forum" then
	Response.Redirect "web_showword.asp?id=" & Request("rootid") & "&" & mid(get_param_str(),2)
end if
%>
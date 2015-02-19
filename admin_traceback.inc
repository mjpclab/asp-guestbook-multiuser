<%
if AdminDisplayMode()="book" then
	if Request("type")<>"" and Request("searchtxt")<>"" then
		Response.Redirect "admin_search.asp?user=" &Request("user")& "&page=" & Request("page") & "&type=" & server.URLEncode(Request("type")) & "&searchtxt=" & server.URLEncode(Request("searchtxt"))
	else
		Response.Redirect "admin.asp?user=" &Request("user")& "&page=" & Request("page")
	end if
elseif AdminDisplayMode()="forum" then
	Response.Redirect "admin_showword.asp?id=" & Request("rootid") & "&user=" &Request("user")& "&page=" & Request("page") & "&type=" & server.URLEncode(Request("type")) & "&searchtxt=" & server.URLEncode(Request("searchtxt"))
end if
%>
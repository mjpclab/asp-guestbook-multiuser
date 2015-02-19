<%
function get_param_str
	dim param_str
	param_str=""
	if request("s_adminname")<>"" then param_str=param_str & "&s_adminname=" & server.URLEncode(request("s_adminname"))
	if request("s_name")<>"" then param_str=param_str & "&s_name=" & server.URLEncode(request("s_name"))
	if request("s_title")<>"" then param_str=param_str & "&s_title=" & server.URLEncode(request("s_title"))
	if request("s_article")<>"" then param_str=param_str & "&s_article=" & server.URLEncode(request("s_article"))
	if request("s_email")<>"" then param_str=param_str & "&s_email=" & server.URLEncode(request("s_email"))
	if request("s_qqid")<>"" then param_str=param_str & "&s_qqid=" & server.URLEncode(request("s_qqid"))
	if request("s_msnid")<>"" then param_str=param_str & "&s_msnid=" & server.URLEncode(request("s_msnid"))
	if request("s_homepage")<>"" then param_str=param_str & "&s_homepage=" & server.URLEncode(request("s_homepage"))
	if request("s_ipaddr")<>"" then param_str=param_str & "&s_ipaddr=" & server.URLEncode(request("s_ipaddr"))
	if request("s_originalip")<>"" then param_str=param_str & "&s_originalip=" & server.URLEncode(request("s_originalip"))
	if request("s_reply")<>"" then param_str=param_str & "&s_reply=" & server.URLEncode(request("s_reply"))
	if isnumeric(request("page")) and request("page")<>"" then param_str=param_str & "&page=" & request("page")
	if len(param_str)<>0 then param_str="?" & right(param_str,len(param_str)-1)
	
	'get_param_str=Server.HTMLEncode(param_str)
	get_param_str=param_str
end function

%>
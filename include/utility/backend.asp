<%
function SetTimelessCookie(byval CookieName,byval CookieValue)
	Response.Cookies(CookieName)=CookieValue
	Response.Cookies(CookieName).Expires=CDate("2038-1-18 0:0:0")
	Response.Cookies(CookieName).Path="/"
end function

function FormOrCookie(byval strname)
	if not isempty(Request.Form(strname)) then
		FormOrCookie=Request.Form(strname)
	elseif not isempty(Request.Cookies(strname)) then
		FormOrCookie=Request.Cookies(strname)
	else
		FormOrCookie=""
	end if
end function

function FormOrSession(byval strname)
	if not isempty(Request.Form(strname)) then
		FormOrSession=Request.Form(strname)
	elseif not isempty(Session(strname)) then
		FormOrSession=Session(strname)
	else
		FormOrSession=""
	end if
end function

function addstat(byref fieldname)
	set cna=server.CreateObject("ADODB.Connection")
	set rsa=server.CreateObject("ADODB.Recordset")

	Call CreateConn(cna)
	rsa.Open Replace(Replace(sql_common_getstat,"{0}",fieldname),"{1}",adminid),cna,0,1,1

	if rsa.EOF then
		cna.Execute Replace(Replace(sql_common_initstat,"{0}",DateTimeStr(Now())),"{1}",adminid),,1
	elseif isdate(rsa.Fields("startdate"))=false then
		cna.Execute Replace(Replace(sql_common_updatetime,"{0}",DateTimeStr(Now())),"{1}",adminid),,1
	end if
	rsa.Close
	cna.Execute Replace(Replace(sql_common_addstat,"{0}",fieldname),"{1}",adminid),,1

	cna.Close
	set rsa=nothing
	set cna=nothing
end function

function getvcode(byval codelen)
	dim rvalue
	randomize

	rvalue=cstr(fix(abs(rnd*(10^codelen))))
	if len(rvalue)<codelen then rvalue=string(codelen-len(rvalue),"0") & rvalue

	getvcode=rvalue
end function

function GetDisplayMode(FieldName)
	if isempty(Request.Cookies(FieldName)) or (Request.Cookies(FieldName)<>"book" and Request.Cookies(FieldName)<>"forum") then
		GetDisplayMode=DisplayMode
	else
		GetDisplayMode=Request.Cookies(FieldName)
	end if
end function

function GuestDisplayMode
	GuestDisplayMode=GetDisplayMode(InstanceName & "_DisplayMode_" & ruser)
end function

function AdminDisplayMode
	AdminDisplayMode=GetDisplayMode(InstanceName & "_AdminDisplayMode_" & ruser)
end function

function WebDisplayMode
	WebDisplayMode=GetDisplayMode(InstanceName & "_WebDisplayMode_" & ruser)
end function

function GetRequests
	dim rvalue
	rvalue=""
	if Request.QueryString<>"" then
		for each item in Request.QueryString
			if item<>"mode" and item<>"modeflag" and item<>"rpage" then
				if Request.QueryString(item)<>"" then rvalue=rvalue & "&" & item & "=" & Server.URLEncode(Request.QueryString(item))
			end if
		next
	end if
	if Request.Form<>"" then
		for each item in Request.Form
			if item<>"mode" and item<>"modeflag" and item<>"rpage" then
				if Request.Form(item)<>"" then rvalue=rvalue & "&" & item & "=" & Server.URLEncode(Request.Form(item))
			end if
		next
	end if
	GetRequests=HtmlEncode(rvalue)
end function

Function DateTimeStr(theTime)
	dim t
	t=CDate(theTime)
	DateTimeStr = Year(t) & "-" & Month(t) & "-" & Day(t) & " " & Hour(t) & ":" & Minute(t) & ":" & Second(t)
End Function

Function ServerTimeToUTC(time)
	ServerTimeToUTC=DateAdd("n",-ServerTimezoneOffset,time)
End Function

Function UTCToDisplayTime(time)
	UTCToDisplayTime=DateAdd("n",DisplayTimezoneOffset,time)
End Function

function geturlpath()
	dim host,url,buffer,port,httpHost
	if LCase(Request.ServerVariables("HTTPS"))="on" then
		host="https://"
	else
		host="http://"
	end if
	httpHost=Request.ServerVariables("HTTP_HOST")
	if httpHost<>"" then
		host=host & httpHost
	else
		buffer=Request.ServerVariables("SERVER_NAME")
		if IsIPv6(buffer) then
			buffer = "[" & buffer & "]"
		end if
		host=host & buffer

		port=Request.ServerVariables("SERVER_PORT")
		if port<>"" and port<>"80" then
			host=host & ":" & port
		end if
	end if
	host=host & Request.ServerVariables("URL")
	url=left(host,InStrRev(host,"/"))

	geturlpath=url
end function

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
	if param_str<>"" then param_str="?" & right(param_str,len(param_str)-1)

	'get_param_str=HtmlEncode(param_str)
	get_param_str=param_str
end function
%>
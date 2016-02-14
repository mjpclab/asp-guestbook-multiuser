<%
function checkuser(byref conn,byref rec,byval ckform)
dim re
if ckform then
	ruser=Request.Form("user")
else
	ruser=request("user")
end if

if ruser="" then
	conn.Close
	set rec=nothing
	set conn=nothing
	response.Redirect "face.asp"
	Response.End
else
	if ruser<>FilterKeyword(ruser) then
		set re=nothing
		conn.Close
		set rec=nothing
		set conn=nothing
		Response.Redirect "face.asp"
		Response.End
	else
		set re=nothing
		rec.Open Replace(sql_common_checkuser,"{0}",ruser),conn,,,1
		if rec.EOF then
			rec.Close
			conn.Close
			set rec=nothing
			set conn=nothing
			Response.Redirect "face.asp"
			Response.End
		else
			adminid=rec.Fields("adminid")
			rec.Close
		end if
	end if
end if
end function
%>
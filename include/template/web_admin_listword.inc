<%
dim guest_txt,guestflag
dim reply_txt,reply_htmlflag
dim js_status
dim param_url

while Not rs.EOF
	web_getadmininfo()
	guestflag=rs("guestflag")
	isauditting=CBool(guestflag AND 16)
	adminname=rs("adminname")
	%>

	<div class="topic"><%
		outerword(rs)
		if CBool(rs("replied") AND 1) and Not ReplyInWord then outeradminreply(rs)
		if Not ReplyInWord then
			Call showWebAdminMessageTools(rs,adminname)
		end if

		if rs.Fields("parent_id")<=0 and Not ReplyInWord then
			set rs1=Server.CreateObject("ADODB.Recordset")
			rs1.Open Replace(Replace(sql_listword_web,"{0}",rs("id")),"{1}",rs("adminid")),cn,0,1,1
			while not rs1.eof
				guestflag=rs1("guestflag")
				isauditting=CBool(guestflag AND 16)
				ishidden=((guestflag AND 40)=8)
				iswhisper=CBool(guestflag AND 32)
				encrypted=((guestflag AND 96)=96)

				outerword(rs1)
				if rs1.Fields("replied") then outeradminreply(rs1)
					Call showWebAdminMessageTools(rs1,adminname)
				rs1.movenext
			wend
			rs1.close
			set rs1=nothing
		end if
		%>
	</div>

	<%rs.MoveNext
wend
%>
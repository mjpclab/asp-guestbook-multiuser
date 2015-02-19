<table id="titlelist" border="1" cellpadding="2" class="generalwindow grid" style="color:<%=TableContentColor%>; table-layout:fixed;">
<tr class="header"><th style="width:25px; text-align:center;"> </th><th style="width:25px;"> </th><th>标题</th><th style="width:100px;">姓名</th><th style="width:120px;">留言时间</th><th style="width:80px;">所属用户</th></tr>
<%
'dim guest_txt,guestflag,isauditting,ishidden,iswhisper,encrypted
'dim reply_txt,reply_htmlflag

for i=1 to TitlesPerPage
	if rs.eof=false and Response.IsClientConnected then
		guestflag=rs("guestflag")
		if clng(guestflag and 16)<>0 then isauditting=true else isauditting=false
		if clng(guestflag and 40)=8 then ishidden=true else ishidden=false
		if clng(guestflag and 32)<>0 then iswhisper=true else iswhisper=false
		if clng(guestflag and 96)=96 then encrypted=true else encrypted=false

		Response.Write Replace("<tr style=""background-color:{0}"">","{0}",TableContentBGC)
		
		Response.Write Replace("<td><input type=""checkbox"" name=""seltodel"" value=""{0}"" /></td>","{0}",rs.Fields("id"))
		
		if isauditting then
			Response.Write "<td style=""text-align:center;""><img class=""forumicon"" src=""image/icon_wait2pass.gif"" alt=""待审核"" /></td>"
		elseif ishidden then
			Response.Write "<td style=""text-align:center;""><img class=""forumicon"" src=""image/icon_hide.gif"" alt=""被隐藏"" /></td>"
		elseif iswhisper then 
			Response.Write "<td style=""text-align:center;""><img class=""forumicon"" src=""image/icon_whisper.gif"" alt=""悄悄话"" /></td>"
		elseif rs.Fields("parent_id")<0 then
			Response.Write "<td style=""text-align:center;""><img class=""forumicon"" src=""image/icon_toplocked.gif"" alt=""置顶"" /></td>"
		else
			Response.Write "<td style=""text-align:center;""> </td>"
		end if
		
		dim title_pattern
		title_pattern="<td nowrap=""nowrap"" style=""overflow:hidden;""><a href=""web_showword.asp?id={0}{1}"">{2}</a></td>"
		if Trim(rs.Fields("title"))="" then
			Response.Write Replace(Replace(Replace(title_pattern,"{0}",rs.Fields("id")),"{1}",GetRequests()),"{2}","(无标题)")
		else
			Response.Write Replace(Replace(Replace(title_pattern,"{0}",rs.Fields("id")),"{1}",GetRequests()),"{2}",rs.Fields("title"))
		end if
		
		Response.Write "<td nowrap=""nowrap"" style=""overflow:hidden;"">" & rs.Fields("name") & "</td>"
		Response.Write "<td nowrap=""nowrap"" style=""overflow:hidden;"">" & rs.Fields("logdate") & "</td>"
		Response.Write Replace("<td nowrap=""nowrap"" style=""overflow:hidden;""><a href=""web_userinfo.asp?user={0}"" target=""_blank"">{0}</a></td>","{0}",rs.Fields("adminname"))
		
		rs.MoveNext
	else
		exit for
	end if
next
%>
</table>
<!-- #include file="rowlight.inc" -->
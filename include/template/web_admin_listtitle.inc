<table class="topic-list">
<thead>
<tr>
	<th class="select"><input type="checkbox" name="seltodel" class="seltodel checkbox" value=""/></th><th class="icon"></th><th class="title">����</th><th class="name">����</th><th class="date">����ʱ��</th><th class="user">�����û�</th>
</tr>
</thead>
<tbody>
<%
'dim guest_txt,guestflag,isauditting,ishidden,iswhisper,encrypted
'dim reply_txt,reply_htmlflag

while Not rs.eof
	guestflag=rs("guestflag")
	isauditting=CBool(guestflag AND 16)
	ishidden=((guestflag AND 40)=8)
	iswhisper=CBool(guestflag AND 32)
	encrypted=((guestflag AND 96)=96)
	%>
	<tr class="message<%if isauditting then%> auditting-message<%end if%>">
		<td class="select">
			<input type="checkbox" name="seltodel" class="seltodel checkbox" value="<%=rs.Fields("id")%>" />
		</td>
		<td class="icon">
			<%if isauditting then%><img src="asset/image/icon_wait2pass.gif" title="�����" alt="�����"/>
			<%elseif ishidden then%><img src="asset/image/icon_hide.gif" title="������" alt="������"/>
			<%elseif iswhisper then%><img src="asset/image/icon_whisper.gif" title="���Ļ�" alt="���Ļ�"/>
			<%elseif rs.Fields("parent_id")<0 then%><img src="asset/image/icon_toplocked.gif" title="�ö�" alt="�ö�"/>
			<%end if%>
		</td>
		<td class="title">
			<%if Trim(rs.Fields("title"))="" then%><a href="showword.asp?id=<%=rs.Fields("id") & GetRequests()%>">(�ޱ���)</a>
			<%else%><a href="web_showword.asp?id=<%=rs.Fields("id") & GetRequests()%>" title="<%=rs.Fields("title")%>"><%=rs.Fields("title")%></a>
			<%end if%>
		</td>
		<td class="name"><%=rs.Fields("name")%></td>
		<td class="date"><%=UTCToDisplayTime(rs.Fields("logdate"))%></td>
		<td class="user"><a href="web_userinfo.asp?user=<%=rs.Fields("adminname")%>" target="_blank"><%=rs.Fields("adminname")%></a></td>
	</tr>
	<%rs.MoveNext
wend
%>
</tbody>
</table>
<script type="text/javascript" src="asset/js/table-select.js"></script>

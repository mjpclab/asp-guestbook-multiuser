<table class="topic-list">
<thead>
<tr>
	<th class="icon"></th><th class="title">����</th><th class="name">����</th><th class="date">����ʱ��</th>
</tr>
</thead>
<tbody>
<%
'dim guest_txt,guestflag,isauditting,ishidden,iswhisper,encrypted
'dim reply_txt,reply_htmlflag
n_i=0

while rs.eof=false
	guestflag=rs("guestflag")
	isauditting=(clng(guestflag and 16)<>0)
	ishidden=(clng(guestflag and 40)=8)
	iswhisper=(clng(guestflag and 32)<>0)
	encrypted=(clng(guestflag and 96)=96)

	n_i=n_i+1
	%>
	<tr>
		<td class="icon">
			<a name="a<%=n_i%>" id="a<%=n_i%>"></a>
			<%if isauditting then%><img src="image/icon_wait2pass.gif" title="�����" alt="�����"/>
			<%elseif ishidden then%><img src="image/icon_hide.gif" title="������" alt="������"/>
			<%elseif iswhisper then%><img src="image/icon_whisper.gif" title="���Ļ�" alt="���Ļ�"/>
			<%elseif rs.Fields("parent_id")<0 then%><img src="image/icon_toplocked.gif" title="�ö�" alt="�ö�"/>
			<%end if%>
		</td>
		<td class="title">
			<%if isauditting then%><a href="showword.asp?user=<%=ruser%>&id=<%=rs.Fields("id")%>">(���Դ����...)</a>
			<%elseif ishidden then%><a href="showword.asp?user=<%=ruser%>&id=<%=rs.Fields("id")%>">(���Ա�����Ա����...)</a>
			<%elseif iswhisper then%><a href="showword.asp?user=<%=ruser%>&id=<%=rs.Fields("id")%>">(�����������Ļ�...)</a>
			<%elseif Trim(rs.Fields("title"))="" then%><a href="showword.asp?user=<%=ruser%>&id=<%=rs.Fields("id")%>">(�ޱ���)</a>
			<%else%><a href="showword.asp?user=<%=ruser%>&id=<%=rs.Fields("id")%>" title="<%=rs.Fields("title")%>"><%=rs.Fields("title")%></a>
			<%end if%>
		</td>
		<td class="name"><%if isauditting=false then Response.Write rs.Fields("name")%></td>
		<td class="date"><%=rs.Fields("logdate")%></td>
	</tr>
	<%rs.MoveNext
wend
%>
</tbody>
</table>
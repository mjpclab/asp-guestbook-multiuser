<%
param_url="?rootid=" & p_rs.Fields("root_id") & "&id=" & p_rs("id")
param_temp=get_param_str()
if param_temp<>"" then param_temp="&" & mid(param_temp,2)
param_url=param_url & param_temp
%>
<div class="admin-tools">
	<span class="tool"><input type="checkbox" name="seltodel" id="c<%=p_rs("id")%>" value="<%=p_rs("id")%>" /><label for="c<%=p_rs("id")%>">(ѡ��)</label></span>
	<span class="tool">�����û���<a href="web_userinfo.asp?user=<%=p_rs("adminname")%>" target="_blank"><%=p_rs("adminname")%></a></span>
	<a class="tool" href="web_searchdel.asp<%=param_url%>" title="ɾ������"<%if DelTip=true then Response.Write " onclick=""return confirm('ȷʵҪɾ��������');"""%>><img src="image/icon_del.gif"/>[ɾ��]</a>
	<%if p_rs("replied")=true then%><a class="tool" href="web_searchdelreply.asp?id=<%=p_rs("id")%><%if isnumeric(Request.QueryString("page")) and Request.QueryString("page")<>"" then response.write "&page=" & Request.QueryString("page")%><%=param_str%>" title="ɾ���ظ�"<%if DelReTip=true then Response.Write " onclick=""return confirm('ȷʵҪɾ���ظ���')"""%>><img border="0" src="image/icon_delreply.gif" />[ɾ���ظ�]</a><%end if%>
	<%if clng(guestflag and 16)<>0 then%><span class="tool"><img src="image/icon_wait2pass.gif">�����</span><%end if%>
	<%if clng(guestflag and 32)<>0 then%><span class="tool"><img src="image/icon_whisper.gif">���Ļ�<%if clng(guestflag and 64)<>0 then response.write ",�Ѽ���"%></span><%end if%>
	<%if p_rs.Fields("parent_id")<0 then%><span class="tool"><img src="image/icon_toplocked.gif" />�������ö�</span><%end if%>
	<%if clng(guestflag and 512)<>0 and p_rs.Fields("parent_id")<=0 then%><span class="tool"><img src="image/icon_lockreply.gif" />�ظ�������</span><%end if%>
	<%if clng(guestflag and 256)<>0 then%><span class="tool"><img src="image/icon_hidecontact.gif" />��ϵ������</span><%end if%>
	<%if clng(guestflag and 40)=8 then%><span class="tool"><img src="image/icon_hide.gif" />����������</span><%end if%>
	<%if clng(guestflag and 128)<>0 and MailReplyInform=true then%><span class="tool"><img src="image/icon_mail.gif" />�ظ�֪ͨ</span><%end if%>
</div>
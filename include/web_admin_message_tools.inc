<%
param_url="?rootid=" & p_rs.Fields("root_id") & "&id=" & p_rs("id")
param_temp=get_param_str()
if param_temp<>"" then param_temp="&" & mid(param_temp,2)
param_url=param_url & param_temp
%>
<div class="admin-message-tools">
	<span class="tool"><input type="checkbox" name="seltodel" id="c<%=p_rs("id")%>" value="<%=p_rs("id")%>" /><label for="c<%=p_rs("id")%>">(选定)</label></span>
	<span class="tool">所属用户：<a href="web_userinfo.asp?user=<%=adminname%>" target="_blank"><%=adminname%></a></span>
	<a class="tool" href="web_searchdel.asp<%=param_url%>" title="删除留言"<%if DelTip=true then Response.Write " onclick=""return confirm('确实要删除留言吗？');"""%>><img src="image/icon_del.gif"/>[删除]</a>
	<%if p_rs("replied")=true then%><a class="tool" href="web_searchdelreply.asp?id=<%=p_rs("id")%><%if isnumeric(Request.QueryString("page")) and Request.QueryString("page")<>"" then response.write "&page=" & Request.QueryString("page")%><%=param_str%>" title="删除回复"<%if DelReTip=true then Response.Write " onclick=""return confirm('确实要删除回复吗？')"""%>><img border="0" src="image/icon_delreply.gif" />[删除回复]</a><%end if%>
	<%if clng(guestflag and 16)<>0 then%><span class="tool"><img src="image/icon_wait2pass.gif">待审核</span><%end if%>
	<%if clng(guestflag and 32)<>0 then%><span class="tool"><img src="image/icon_whisper.gif">悄悄话<%if clng(guestflag and 64)<>0 then response.write ",已加密"%></span><%end if%>
	<%if p_rs.Fields("parent_id")<0 then%><span class="tool"><img src="image/icon_toplocked.gif" />留言已置顶</span><%end if%>
	<%if clng(guestflag and 512)<>0 and p_rs.Fields("parent_id")<=0 then%><span class="tool"><img src="image/icon_lockreply.gif" />回复已锁定</span><%end if%>
	<%if clng(guestflag and 256)<>0 then%><span class="tool"><img src="image/icon_hidecontact.gif" />联系已隐藏</span><%end if%>
	<%if clng(guestflag and 40)=8 then%><span class="tool"><img src="image/icon_hide.gif" />内容已隐藏</span><%end if%>
	<%if clng(guestflag and 128)<>0 and MailReplyInform=true then%><span class="tool"><img src="image/icon_mail.gif" />回复通知</span><%end if%>
</div>

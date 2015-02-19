<%
param_url="?rootid=" & p_rs.Fields("root_id") & "&id=" & p_rs("id")
param_temp=get_param_str()
if param_temp<>"" then param_temp="&" & mid(param_temp,2)
param_url=param_url & param_temp
%>
<tr>
	<td colspan="2" style="width:100%; border:solid <%=TableBorderColor%>; border-width:1px 0px 1px 0px; padding:3px; color:<%=TableContentColor%>; background-color:<%=TableContentBGC%>;">
		<input type="checkbox" name="seltodel" id="c<%=p_rs("id")%>" value="<%=p_rs("id")%>" /><label for="c<%=p_rs("id")%>">(选定)</label>
		<span style="font-weight:bold">所属用户：<a href="web_userinfo.asp?user=<%=p_rs("adminname")%>" target="_blank"><%=p_rs("adminname")%></a></span> 
		
		<a href="web_searchdel.asp<%=param_url%>" title="删除留言"<%if DelTip=true then Response.Write " onclick=""return confirm('确实要删除留言吗？');"""%>><img alt="" src="image/icon_del.gif" class="imgicon" />[删除]</a> 
		<%if p_rs("replied")=true then%><a href="web_searchdelreply.asp?id=<%=p_rs("id")%><%if isnumeric(Request.QueryString("page")) and Request.QueryString("page")<>"" then response.write "&amp;page=" & Request.QueryString("page")%><%=param_str%>" title="删除回复"<%if DelReTip=true then Response.Write " onclick=""return confirm('确实要删除回复吗？')"""%>><img border="0" src="image/icon_delreply.gif" class="imgicon" />[删除回复]</a> <%end if%>
		<%if clng(guestflag and 16)<>0 then%><img alt="" src="image/icon_wait2pass.gif" class="imgicon">待审核 <%end if%>
		<%if clng(guestflag and 32)<>0 then%><img alt="" src="image/icon_whisper.gif" class="imgicon">悄悄话<%if clng(guestflag and 64)<>0 then response.write ",已加密"%> <%end if%>
		<%if p_rs.Fields("parent_id")<0 then%><img alt="" src="image/icon_toplocked.gif" class="imgicon" />留言已置顶  <%end if%>
		<%if clng(guestflag and 512)<>0 and p_rs.Fields("parent_id")<=0 then%><img alt="" src="image/icon_lockreply.gif" class="imgicon" />回复已锁定  <%end if%>
		<%if clng(guestflag and 256)<>0 then%><img alt="" src="image/icon_hidecontact.gif" class="imgicon" />联系已隐藏 <%end if%>
		<%if clng(guestflag and 40)=8 then%><img alt="" src="image/icon_hide.gif" class="imgicon" />内容已隐藏 <%end if%>
		<%if clng(guestflag and 128)<>0 and MailReplyInform=true then%><img alt="" src="image/icon_mail.gif" class="imgicon" />回复通知 <%end if%>
	</td>
</tr>

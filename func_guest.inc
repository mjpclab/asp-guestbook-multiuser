<table class="generalwindow" style="border-width:0px;">
	<tr><td style="text-align:left;">
		<%if StatusWrite=true then%><a href="leaveword.asp?user=<%=request("user")%>" title="签写留言"><img alt="" src="image/icon_write.gif" style="border-width:0px" />签写留言</a><%end if%>
		 
		<%if RPage<>"" then%>
			<%if GuestDisplayMode="book" then%>
				<a href="setdisplaymode.asp?mode=forum&amp;modeflag=guest&amp;rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="image/icon_mode_forum.gif" style="border-width:0px;" />切换到论坛模式</a>
			<%elseif GuestDisplayMode="forum" then%>
				<a href="setdisplaymode.asp?mode=book&amp;modeflag=guest&amp;rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="image/icon_mode_book.gif" style="border-width:0px;" />切换到留言本模式</a>
			<%end if%>
		<%end if%>
	</td><td style="text-align:right;">
		<span style="color:<%=LinkNormal%>;">| <a href="reg.asp">创建属于您的留言本</a> | <a href="admin.asp?user=<%=request("user")%>">管理</a></span>
	</td></tr>
</table>
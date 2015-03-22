<div class="guest-functions">
	<div class="main">
		<%if StatusWrite=true then%><a class="function" href="leaveword.asp?user=<%=ruser%>" title="签写留言"><img src="image/icon_write.gif"/>签写留言</a><%end if%>
		<%if RPage<>"" then%>
			<%if GuestDisplayMode="book" then%>
				<a class="function" href="setdisplaymode.asp?mode=forum&modeflag=guest&rpage=<%=RPage%><%=GetRequests()%>"><img src="image/icon_mode_forum.gif" />切换到论坛模式</a>
			<%elseif GuestDisplayMode="forum" then%>
				<a class="function" href="setdisplaymode.asp?mode=book&modeflag=guest&rpage=<%=RPage%><%=GetRequests()%>"><img src="image/icon_mode_book.gif" />切换到留言本模式</a>
			<%end if%>
		<%end if%>
	</div>
	<div class="aside">
		<a class="function" href="reg.asp">创建属于您的留言本</a>
		<a class="function" href="admin.asp?user=<%=ruser%>">管理</a>
	</div>
</div>

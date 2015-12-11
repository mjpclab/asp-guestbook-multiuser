<div class="guest-functions">
	<div class="main">
		<a class="function function-multi-del" href="javascript:for(var i=0;i<=form7.elements.seltodel.length-1;i++)if(form7.elements.seltodel[i].checked){<%if DelSelTip=true then Response.Write "if (confirm('确实要删除选定留言吗？')==true)"%>form7.submit();break;}else if(i==form7.elements.seltodel.length-1)alert('请先选定要删除的留言。');"><img src="image/icon_mdel.gif"/>删除选定留言</a>
		<%if RPage<>"" then%>
			<%if WebDisplayMode="book" then%>
				<a class="function" href="web_setdisplaymode.asp?mode=forum&rpage=<%=RPage%><%=GetRequests()%>"><img src="image/icon_mode_forum.gif"/>切换到标题模式</a>
			<%elseif WebDisplayMode="forum" then%>
				<a class="function" href="web_setdisplaymode.asp?mode=book&&rpage=<%=RPage%><%=GetRequests()%>"><img src="image/icon_mode_book.gif"/>切换到完整模式</a>
			<%end if%>
		<%end if%>
	</div>
</div>

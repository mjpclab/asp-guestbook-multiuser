<span class="generalwindow noborder" style="text-align:left; display:block;">
	<a href="javascript:for(var i=0;i&lt;=form7.elements.length-1;i++)if(form7.elements[i].name=='seltodel' && form7.elements[i].checked){<%if DelSelTip=true then Response.Write "if (confirm('确实要删除选定留言吗？')==true)"%>form7.submit();break;}else if(i==form7.elements.length-1)alert('请先选定要删除的留言。');" onmouseover="return true;"><img alt="" src="image/icon_mdel.gif" style="border-width:0px;" />删除选定留言</a>  
	<%if RPage<>"" then%>
		<%if WebDisplayMode="book" then%>
			<a href="setdisplaymode.asp?mode=forum&amp;modeflag=web&amp;rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="image/icon_mode_forum.gif" style="border-width:0px;" />切换到论坛模式</a>
		<%elseif WebDisplayMode="forum" then%>
			<a href="setdisplaymode.asp?mode=book&amp;modeflag=web&amp;rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="image/icon_mode_book.gif" style="border-width:0px;" />切换到留言本模式</a>
		<%end if%>
	<%end if%>
</span>

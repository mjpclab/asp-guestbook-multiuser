<div class="guest-functions">
	<div class="main">
		<a class="function function-multi-del" href="javascript:for(var i=0;i<=form7.elements.seltodel.length-1;i++)if(form7.elements.seltodel[i].checked){<%if DelSelTip=true then Response.Write "if (confirm('ȷʵҪɾ��ѡ��������')==true)"%>form7.submit();break;}else if(i==form7.elements.seltodel.length-1)alert('����ѡ��Ҫɾ�������ԡ�');"><img src="image/icon_mdel.gif"/>ɾ��ѡ������</a>
		<a class="function function-multi-pass" href="javascript:for(var i=0;i<=form7.elements.seltodel.length-1;i++)if(form7.elements.seltodel[i].checked){<%if PassSelTip=true then Response.Write "if (confirm('ȷʵҪͨ�����ѡ��������')==true)"%>{form7.action='admin_mpassaudit.asp';form7.submit();}break;}else if(i==form7.elements.seltodel.length-1)alert('����ѡ��Ҫͨ����˵����ԡ�');"><img src="image/icon_mpass.gif"/>ͨ�����ѡ������</a>
		<%if RPage<>"" then%>
			<%if AdminDisplayMode="book" then%>
				<a class="function" href="setdisplaymode.asp?mode=forum&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img src="image/icon_mode_forum.gif"/>�л�������ģʽ</a>
			<%elseif AdminDisplayMode="forum" then%>
				<a class="function" href="setdisplaymode.asp?mode=book&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img src="image/icon_mode_book.gif"/>�л�������ģʽ</a>
			<%end if%>
		<%end if%>
	</div>
</div>
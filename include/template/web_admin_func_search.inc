<div class="guest-functions">
	<div class="main">
		<a class="function function-multi-del" href="javascript:void 0;"><img src="asset/image/icon_mdel.gif"/>删除选定留言</a>
		<%if RPage<>"" then%>
			<%if WebDisplayMode="book" then%>
				<a class="function" href="web_setdisplaymode.asp?mode=forum&rpage=<%=RPage%><%=GetRequests()%>"><img src="asset/image/icon_mode_forum.gif"/>切换到标题模式</a>
			<%elseif WebDisplayMode="forum" then%>
				<a class="function" href="web_setdisplaymode.asp?mode=book&&rpage=<%=RPage%><%=GetRequests()%>"><img src="asset/image/icon_mode_book.gif"/>切换到完整模式</a>
			<%end if%>
		<%end if%>
	</div>
</div>
<script type="text/javascript">
	(function multiDelConfirm(){
		var $multiDelLink = $('.function-multi-del:last');
		$multiDelLink.click(function () {
			var delSelTip = <%=LCase(CStr(DelSelTip))%>;
			var $selected = $('input.seltodel:checked');
			if (!$selected.length) {
				alert('请先选定要删除的留言。');
			} else if (!delSelTip || confirm('确实要删除选定留言吗？')) {
				form7.submit();
			}
		});
	}());
</script>
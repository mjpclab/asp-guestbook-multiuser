<div class="region">
	<h3 class="title">用户管理员控制中心</h3>
	<div class="content">
		<table>
			<tr>
				<td><a href="admin.asp?user=<%=ruser%>">管理首页</a></td>
				<td><a href="admin_search.asp?user=<%=ruser%>">搜索留言</a></td>
				<td><a href="admin_advdel.asp?user=<%=ruser%>">高级删除</a></td>
				<td><a href="admin_setbulletin.asp?user=<%=ruser%>">发布置顶公告</a></td>
				<td><a href="admin_chpass.asp?user=<%=ruser%>">修改密码/问题</a></td>
			</tr>
			<tr>
				<td><a href="admin_setinfo.asp?user=<%=ruser%>">修改版主资料</a></td>
				<td><a href="admin_callgen.asp?user=<%=ruser%>">生成调用代码</a></td>
				<td><a href="admin_config.asp?user=<%=ruser%>&page=1">留言本参数</a></td>
				<td><a href="admin_ipconfig.asp?user=<%=ruser%>">IP屏蔽策略</a></td>
				<td><a href="admin_filter.asp?user=<%=ruser%>">内容过滤策略</a></td>
			</tr>
			<tr>
				<td><a href="admin_floodconfig.asp?user=<%=ruser%>">防灌水策略</a></td>
				<td><a href="admin_stats.asp?user=<%=ruser%>">访问统计</a></td>
				<td><a href="admin_logoff.asp?user=<%=ruser%>">退出管理</a></td>
				<td></td>
				<td></td>
			</tr>
		</table>
	</div>
</div>

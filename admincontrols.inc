<div class="region">
	<h3 class="title">用户管理员控制中心</h3>
	<div class="content">
		<table>
			<tr>
				<td><a href="admin.asp?user=<%=request("user")%>">管理首页</a></td>
				<td><a href="admin_search.asp?user=<%=request("user")%>">搜索留言</a></td>
				<td><a href="admin_advdel.asp?user=<%=request("user")%>">高级删除</a></td>
				<td><a href="admin_setbulletin.asp?user=<%=request("user")%>">发布置顶公告</a></td>
				<td><a href="admin_chpass.asp?user=<%=request("user")%>">修改密码/问题</a></td>
			</tr>
			<tr>
				<td><a href="admin_setinfo.asp?user=<%=request("user")%>">修改版主资料</a></td>
				<td><a href="admin_callgen.asp?user=<%=request("user")%>">生成调用代码</a></td>
				<td><a href="admin_config.asp?user=<%=request("user")%>&page=1">留言本参数</a></td>
				<td><a href="admin_ipconfig.asp?user=<%=request("user")%>">IP屏蔽策略</a></td>
				<td><a href="admin_filter.asp?user=<%=request("user")%>">内容过滤策略</a></td>
			</tr>
			<tr>
				<td><a href="admin_floodconfig.asp?user=<%=request("user")%>">防灌水策略</a></td>
				<td><a href="admin_stats.asp?user=<%=request("user")%>">访问统计</a></td>
				<td><a href="admin_logoff.asp?user=<%=request("user")%>">退出管理</a></td>
				<td></td>
				<td></td>
			</tr>
		</table>
	</div>
</div>

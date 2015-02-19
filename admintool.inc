<table border="1" cellpadding="2" class="generalwindow">
	<tr>
		<td class="centertitle">用户管理员工具</td>
	</tr>
	<tr>
		<td style="<%=CssBgColor(TableContentBGC)%>">
			<table border="1" cellpadding="0" class="grid" style="text-align:center;">
				<tr style="cursor:pointer">
					<td style="width:14%" onclick="window.location='admin.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin.asp?user=<%=request("user")%>">管理首页</a></td>
					<td style="width:14%" onclick="window.location='admin_search.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_search.asp?user=<%=request("user")%>">搜索留言</a></td>
					<td style="width:15%" onclick="window.location='admin_advdel.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_advdel.asp?user=<%=request("user")%>">高级删除</a></td>
					<td style="width:15%" onclick="window.location='admin_setbulletin.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_setbulletin.asp?user=<%=request("user")%>">发布置顶公告</a></td>
					<td style="width:14%" onclick="window.location='admin_chpass.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_chpass.asp?user=<%=request("user")%>">修改密码/问题</a></td>
					<td style="width:14%" onclick="window.location='admin_setinfo.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_setinfo.asp?user=<%=request("user")%>">修改版主资料</a></td>
					<td style="width:14%" onclick="window.location='admin_callgen.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_callgen.asp?user=<%=request("user")%>">生成调用代码</a></td>
				</tr>
				<tr style="cursor:pointer">
					<td style="width:14%" onclick="window.location='admin_config.asp?user=<%=request("user")%>&amp;page=1'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_config.asp?user=<%=request("user")%>&amp;page=1">留言本参数</a></td>
					<td style="width:14%" onclick="window.location='admin_ipconfig.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_ipconfig.asp?user=<%=request("user")%>">IP屏蔽策略</a></td>
					<td style="width:15%" onclick="window.location='admin_filter.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_filter.asp?user=<%=request("user")%>">内容过滤策略</a></td>
					<td style="width:15%" onclick="window.location='admin_floodconfig.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_floodconfig.asp?user=<%=request("user")%>">防灌水策略</a></td>
					<td style="width:14%" onclick="window.location='admin_stats.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_stats.asp?user=<%=request("user")%>">访问统计</a></td>
					<td style="width:14%" onclick="window.location='admin_logoff.asp?user=<%=request("user")%>'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_logoff.asp?user=<%=request("user")%>">退出管理</a></td>
					<td>&nbsp;</td>
				</tr>
			</table>
		</td>
	</tr>
</table>

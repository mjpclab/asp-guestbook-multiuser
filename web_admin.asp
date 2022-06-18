<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/web.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/const.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/web_admin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="include/template/dtd.inc" -->
<html lang="zh-CN">
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster管理中心</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<%Call WebInitHeaderData("","Webmaster管理中心","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<h3>注册用户列表：</h3>
	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)

	Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
	get_divided_page cn,rs,sql_pk_supervisor,sql_webadmin_words_count,sql_webadmin_words_query,"regdate DEC",Request.QueryString("page"),TitlesPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage

	if ItemsCount>0 then
	%>
		<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"[用户列表分页，共" &PagesCount& "页，" &ItemsCount& "个用户]",""%>
		<form method="post" action="web_deluser.asp" name="frm_user" onsubmit="for(var i=0;i<=elements.users.length-1;i++)if(elements.users[i].checked){if(confirm('警告！删除的用户和数据将不能恢复！\n确实要执行删除操作吗？'))return confirm('请再次确认是否要删除用户？');else return false;}alert('请先选择要删除的用户。');return false;">

			<input type="hidden" name="source" value="1">
			<%if isnumeric(Request.QueryString("page")) and Request.QueryString("page")<>"" then%><input type="hidden" name="arguments" value="page=<%=Request.QueryString("page")%>"><%end if%>

			<table id="titlelist" class="topic-list">
			<thead>
				<tr><th class="select"><input type="checkbox" name="users" class="users checkbox"/></th><th>用户名</th><th>昵称</th><th>注册日期</th><th>最后登录日期</th><th>打开留言本</th></tr>
			</thead>
			<tbody>
			<%while not rs.EOF%>
				<tr><td class="select"><input type="checkbox" name="users" class="users checkbox" value="<%=rs("adminname")%>" /></td><td><a href="web_userinfo.asp?user=<%=rs("adminname")%>" target="_blank"><%=rs("adminname")%></a></td><td><%=rs("name")%></td><td><%=UTCToDisplayTime(rs("regdate"))%></td><td><%=UTCToDisplayTime(rs("lastlogin"))%></td><td><a href="index.asp?user=<%=rs("adminname")%>" target="_blank">打开留言本</a></td></tr>
			<%rs.MoveNext : wend%>
			</tbody>
			</table>

			<%if ItemsCount>0 then%>
			<div class="guest-functions">
			<div class="main">
				<input type="submit" value="删除选定用户" />
			</div>
			</div>
			<%end if%>
		</form>
	<%
	rs.Close
	end  if
	cn.Close : set rs=nothing : set cn=nothing
	%>

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"[用户列表分页，共" &PagesCount& "页，" &ItemsCount& "个用户]",""%>

	<!-- #include file="include/template/web_admin_searchuserbox.inc" -->
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
<script type="text/javascript" src="asset/js/jquery-1.x-min.js"></script>
<script type="text/javascript" src="asset/js/table-select.js"></script>
</body>
</html>
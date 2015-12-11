<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="include/utility/book.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 错误</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"错误"%>
	<div class="guest-functions">
		<div class="aside">
			<a class="function" href="reg.asp">申请留言本</a>
			<a class="function" href="admin.asp?user=<%=ruser%>">管理</a>
		</div>
	</div>

	<%
	dim errmsg
	select case Request("number")
	case "1"
		errmsg="对不起，您的访问已被禁止。"
	case "2"
		errmsg="抱歉，留言本已关闭，请稍候再试。"
	case "3"
		errmsg="抱歉，留言权限已关闭，请稍候再试。"
	case "4"
		errmsg="对不起，您的留言中含有禁止出现的内容。"
	case "5"
		errmsg="抱歉，搜索权限已关闭，请稍候再试。"
	case "6"
		errmsg="对不起，您的发言速度太快了，请休息一下。"
	case "7"
		errmsg="对不起，请不要发送重复内容。"
	case else
		errmsg="未知错误，请联系管理员。"
	end select
	Response.Write "<br/><br/><span style=""width:100%; text-align:left;"">　　■　" &errmsg& "</span><br/><br/><br/>"
	%>
</div>

<!-- #include file="include/template/footer.inc" -->
</body>
</html>
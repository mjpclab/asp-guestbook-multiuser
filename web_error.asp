<%
sub WebErrorPage(errorCode)
	Response.Status="403 Forbidden"

	dim errmsg
	select case errorCode
	case 1
		errmsg="对不起，留言本创建功能已关闭，请稍后再试。"
	case 2
		errmsg="对不起，找回密码功能已关闭，请稍后再试。"
	case 3
		errmsg="对不起，管理维护功能已关闭，请稍后再试。"
	case 4
		errmsg="对不起，您的访问已被禁止，请稍后再试。"
	case 5
		errmsg="对不起，自删帐号功能已关闭，请稍后再试。"
	case else
		errmsg="未知错误，请联系管理员。"
	end select
	%>
	<!-- #include file="include/template/dtd.inc" -->
	<html>
	<head>
		<!-- #include file="include/template/metatag.inc" -->
		<title><%=web_BookName%> 错误</title>
		<!-- #include file="inc_web_admin_stylesheet.asp" -->
	</head>

	<body>

	<%Call WebInitHeaderData("","错误","","")%><!-- #include file="include/template/header.inc" -->
	<div id="outerborder" class="outerborder">
		<div id="mainborder" class="mainborder">
		<!-- #include file="include/template/web_guest_func.inc" -->
		<p style="margin-bottom: 3em;">　　■　<%=errmsg%></p>
		</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
	</body>
	</html>
	<%
end sub
%>

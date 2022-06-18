<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/web.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
end if
%>

<!-- #include file="include/template/dtd.inc" -->
<html lang="zh-CN">
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%></title>
	<!-- #include file="inc_stylesheet.asp" -->
	<link rel="stylesheet" type="text/css" href="asset/css/face_style.css"/>
</head>

<body>

<div id="outerborder" class="outerborder">

<%Call WebInitHeaderData("","","","")%><!-- #include file="include/template/header.inc" -->
<div id="mainborder" class="mainborder">
<%
set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

dim sys_bul_flag
sys_bul_flag=16
%>
<!-- #include file="include/template/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/template/web_guest_func.inc" -->

<div class="main">
	<div class="image">
		<img src="asset/image/face_preview_1.png"/>
		<img src="asset/image/face_preview_2.png"/>
		<img src="asset/image/face_preview_3.png"/>
	</div>
	<div class="content">
		<h3>特性介绍</h3>

		<h4>留言支持UBB、HTML，由版主决定开启或关闭</h4>
		<p>通过UBB或HTML能为版面增加不少效果，有时甚至能使表述更明确。但有鉴于HTML代码带来的风险性，不建议版主开放HTML功能，而采用相对较安全、但功能偏于简单的UBB标记。</p>

		<h4>可随时开启或关闭的访客搜索</h4>
		<p>访客不会再为找不到以前的留言而犯愁，可通过访客姓名、标题、留言内容、版主回复搜索。</p>

		<h4>多种配色方案供选择</h4>
		<p>版主可以自选留言本配色方案，以便和网页色调保持一致。</p>

		<h4>IP屏蔽策略将恶意人士拒之门外</h4>
		<p>通过IP屏蔽策略，可以将图谋不轨之人拒之门外，也可以只允许特定IP的访问，所要做的只是启用策略并输入IP段而已。</p>

		<h4>内容过滤策略：更灵活的安全措施</h4>
		<p>如果说IP屏蔽策略适用于完全屏蔽、长期生效的固定IP，那么内容过滤策略将会是更灵活的解决方案。该策略允许使用通配符或正则表达式来匹配规律性的字符串，并且版主可以根据情况选择替换字符串或者直接拒绝留言。</p>

		<h4>移动设备屏幕自适应</h4>
		<p>针对不同尺寸的屏幕，都能自动调整至最佳布局。</p>

		<h4>其它功能</h4>
		<p>留言审核、悄悄话、邮件通知等。</p>
	</div>
</div>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
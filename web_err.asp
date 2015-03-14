<!-- #include file="webconfig.asp" -->
<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> 错误</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<link rel="stylesheet" type="text/css" href="css/web_adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
	<!-- #include file="css/web_adminstyle.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" class="booktitle">首页</a> &gt;&gt; 错误</div>
</div>

<!-- #include file="func_web.inc" -->

<%
dim errmsg
select case Request("number")
case "1"
	errmsg="对不起，留言本创建功能已关闭，请稍候再试。"
case "2"
	errmsg="对不起，找回密码功能已关闭，请稍候再试。"
case "3"
	errmsg="对不起，管理维护功能已关闭，请稍候再试。"
case "4"
	errmsg="对不起，您的访问已被禁止，请稍候再试。"
case "5"
	errmsg="对不起，自删帐号功能已关闭，请稍候再试。"
case else
	errmsg="未知错误，请联系管理员。"
end select
Response.Write "<br/><br/><p>　　■　" &errmsg& "</p><br/><br/><br/>"
%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
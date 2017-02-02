<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/web.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/web_admin_userinfo.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%Response.Expires=-1%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster管理中心 查看用户信息</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
checkuser cn,rs,false

%>

<div id="outerborder" class="outerborder">

<%Call WebInitHeaderData("","Webmaster管理中心","","")%><!-- #include file="include/template/header.inc" -->

<div id="mainborder" class="mainborder">
<div class="region">
	<h3 class="title">查看用户信息</h3>
	<div class="content">
		<%
		rs.Open Replace(sql_webuserinfo,"{0}",ruser),cn,,,1
		if rs("faceid")>0 then Response.Write "<img src=""asset/face/" & rs("faceid") & ".gif" & """>"
		%>
		<p class="field">
			<a href="index.asp?user=<%=ruser%>" target="_blank">[打开留言本]</a>
			<a href="web_chuserpass.asp?user=<%=ruser%>">[重设用户密码]</a>
			<a href="web_chuserstatus.asp?user=<%=ruser%>">[更改账号状态]</a>
		</p>
		<table cellpadding="5">
			<tr><th scope="row">用户名</th><td><%=rs("adminname")%></td></tr>
			<tr><th scope="row">昵称</th><td><%=rs("name") & ""%></td></tr>
			<tr><th scope="row">E-mail</th><td><%=rs("email") & ""%></td></tr>
			<tr><th scope="row">QQ号</th><td><%=rs("qqid") & ""%></td></tr>
			<tr><th scope="row">Skype</th><td><%=rs("msnid") & ""%></td></tr>
			<tr><th scope="row">主页</th><td><%=rs("homepage") & ""%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_status,"{0}",adminid),cn,,,1
			Dim userstatus
			userstatus=rs.Fields(0)
			%>
			<tr><th scope="row">账号状态</th><td><%if CBool(userstatus AND 1073741824) then Response.Write "禁用" else Response.Write "启用"%></td></tr>
			<tr><th scope="row">账号管理登陆状态</th><td><%if CBool(userstatus AND 536870912) then Response.Write "禁用" else Response.Write "启用"%></td></tr>
			<tr><th scope="row">账号签写留言状态</th><td><%if CBool(userstatus AND 268435456) then Response.Write "禁用" else Response.Write "启用"%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_view,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">留言查看次数</th><td><%if Not rs.EOF then Response.Write rs(0) else Response.Write "0"%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_words,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">留言数</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_reply,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">回复数</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_ipv4config,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">自定义IPv4屏蔽策略条数</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_ipv6config,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">自定义IPv6屏蔽策略条数</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_filterconfig,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">自定义内容过滤策略条数</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_reginfo,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">注册日期</th><td><%=UTCToDisplayTime(rs("regdate"))%></td></tr>
			<tr><th scope="row">最后登录日期</th><td><%=UTCToDisplayTime(rs("lastlogin"))%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_logininfo,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">登录次数</th><td><%if not rs.EOF then Response.Write rs(0) else Response.Write "0"%></td></tr>
			<tr><th scope="row">其中登录失败次数</th><td><%if not rs.EOF then Response.Write rs(1) else Response.Write "0"%></td></tr>
		</table>
	</div>
</div>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

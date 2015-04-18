<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster管理中心 查看用户信息</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,false

%>

<div id="outerborder" class="outerborder">

<!-- #include file="web_admintitle.inc" -->


<div class="region">
	<h3 class="title">查看用户信息</h3>
	<div class="content">
		<%
		rs.Open Replace(sql_webuserinfo,"{0}",ruser),cn,,,1
		if rs("faceid")>0 then Response.Write "<img src=""" & FacePath & cstr(rs("faceid")) & ".gif" & """ style=""border:0px;"">"
		%>
		<p class="field">
			<a href="index.asp?user=<%=rs("adminname")%>" target="_blank">[打开留言本]</a>
			<a href="web_chuserpass.asp?user=<%=rs("adminname")%>">[重设用户密码]</a>
		</p>
		<table cellpadding="5">
			<tr><th scope="row">用户名：</th><td><%=rs("adminname")%></td></tr>
			<tr><th scope="row">昵称：</th><td><%=rs("name") & ""%></td></tr>
			<tr><th scope="row">E-mail：</th><td><%=rs("email") & ""%></td></tr>
			<tr><th scope="row">QQ号：</th><td><%=rs("qqid") & ""%></td></tr>
			<tr><th scope="row">Skype：</th><td><%=rs("msnid") & ""%></td></tr>
			<tr><th scope="row">主页：</th><td><%=rs("homepage") & ""%></td></tr>

			<%
			adminid=rs.Fields("adminid")
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_view,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">留言查看次数：</th><td><%if rs.EOF=false then Response.Write rs(0) else Response.Write "0"%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_words,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">留言数：</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_reply,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">回复数：</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_ipv4config,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">自定义IPv4屏蔽策略条数：</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_ipv6config,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">自定义IPv6屏蔽策略条数：</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_filterconfig,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">自定义内容过滤策略条数：</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_reginfo,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">注册日期：</th><td><%=rs("regdate")%></td></tr>
			<tr><th scope="row">最后登录日期：</th><td><%=rs("lastlogin")%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_logininfo,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">登录次数：</th><td><%if not rs.EOF then Response.Write rs(0) else Response.Write "0"%></td></tr>
			<tr><th scope="row">其中登录失败次数：</th><td><%if not rs.EOF then Response.Write rs(1) else Response.Write "0"%></td></tr>
		</table>
	</div>
</div>

<%
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
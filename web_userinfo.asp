<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=web_BookName%> Webmaster管理中心 查看用户信息</title>
	
	<!-- #include file="style.asp" -->
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

<br/>
<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" style="width:100%; border-collapse:collapse;">
	<tr>
		<td class="centertitle">查看用户信息</td>
	</tr>
	<tr>
	<td style="width:100%; color:<%=TableContentColor%>; background-color:<%=TableContentBGC%>;">
		<br/>
		<table  cellpadding="0" cellspacing="0" style="width:100%; border-width:0px;">
			<tr>
			<td style="width:<%=TableLeftWidth%>px; text-align:center; vertical-align:top;">
				<%
				rs.Open Replace(sql_webuserinfo,"{0}",Request.QueryString("user")),cn,,,1
				if rs("faceid")>0 then Response.Write "<img src=""" & FacePath & cstr(rs("faceid")) & ".gif" & """ style=""border:0px;"">"
				%>
				<br/><br/>
				<a href="index.asp?user=<%=rs("adminname")%>" target="_blank">[打开留言本]</a>
				<br/><br/>
				<a href="web_chuserpass.asp?user=<%=rs("adminname")%>">[重设用户密码]</a>
			</td>
			<td style="vertical-align:top;">
				<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" cellspacing="0" style="width:100%; border-collapse:collapse;">
					<tr><td style="width:180px; font-weight:bold;">用户名：</td><td><%=rs("adminname")%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">昵称：</td><td><%=rs("name") & ""%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">E-mail：</td><td><%=rs("email") & ""%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">QQ号：</td><td><%=rs("qqid") & ""%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">MSN：</td><td><%=rs("msnid") & ""%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">主页：</td><td><%=rs("homepage") & ""%></td></tr>
					
					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_view,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">留言查看次数：</td><td><%if rs.EOF=false then Response.Write rs(0) else Response.Write "0"%></td></tr>
					
					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_words,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">留言数：</td><td><%=rs(0)%></td></tr>
					
					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_reply,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">回复数：</td><td><%=rs(0)%></td></tr>
					
					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_ipconfig,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">自定义IP屏蔽策略条数：</td><td><%=rs(0)%></td></tr>

					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_filterconfig,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">自定义内容过滤策略条数：</td><td><%=rs(0)%></td></tr>

					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_reginfo,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">注册日期：</td><td><%=rs("regdate")%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">最后登录日期：</td><td><%=rs("lastlogin")%></td></tr>

					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_logininfo,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">登录次数：</td><td><%if not rs.EOF then Response.Write rs(0) else Response.Write "0"%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">其中登录失败次数：</td><td><%if not rs.EOF then Response.Write rs(1) else Response.Write "0"%></td></tr>
					
				</table>
			</td>
			</tr>
		</table>
		<br/>
	</td>
	</tr>
</table>

<%
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
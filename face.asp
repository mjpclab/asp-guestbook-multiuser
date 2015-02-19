<!-- #include file="webconfig.asp" -->

<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=web_BookName%></title>
	
	<!-- #include file="style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<table cellpadding="2" style="width:100%; border-width:0px; border-collapse:collapse;">
	<tr style="height:60px">
		<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);" nowrap="nowrap">
			<%=web_BookName%>
		</td>
	</tr>
	<tr>
		<td>
			<%
			set cn=server.CreateObject("ADODB.Connection")
			CreateConn cn,dbtype
			
			dim sys_bul_flag
			sys_bul_flag=16
			%>
			<!-- #include file="sysbulletin.inc" -->
			<%cn.Close : set cn=nothing%>
		</td>
	</tr>
	<tr>
		<td style="width:100%;">
		<!-- #include file="func_web.inc" -->
		</td>
	</tr>
	<tr>
		<td style="width:100%; color:<%=TableContentColor%>;">

			<table border="0" width="100%" style="color:<%=TableContentColor%>">
				<tr>
					<td style="width:150px; vertical-align:top;">
						<img src="intro1.gif" style="width:142px; height:159px; border-width:0px;" />
					</td>
					<td style="vertical-align:top;">
						<p style="font-weight:bold">特性介绍：</p>
			
						<p><span style="font-weight:bold">留言支持UBB、HTML，由版主决定开启或关闭</span><br/>
						　　通过UBB或HTML能为版面增加不少效果，有时甚至能使表述更明确。但有鉴于HTML代码带来的风险性，不建议版主开放HTML功能，而采用相对较安全、但功能偏于简单的UBB标记。</p>
						<p><span style="font-weight:bold">可随时开启或关闭的访客搜索</span><br/>
						　　访客不会再为找不到以前的留言而犯愁，可通过访客姓名、标题、留言内容、版主回复搜索。</p>
						<p><span style="font-weight:bold">多种配色方案供选择</span><br/>
						　　版主可以自选留言本配色方案，以便和网页色调保持一致。</p>
						<p><span style="font-weight:bold">验证码有效阻止软件提交</span><br/>
						　　使用验证码可有效防止他人使用软件恶意、重复地提交数据。</p>
						<p><span style="font-weight:bold">IP屏蔽策略将恶意人士拒之门外</span><br/>
						　　通过IP屏蔽策略，可以将图谋不轨之人拒之门外，也可以只允许特定IP的访问，所要做的只是启用策略并输入IP段而已。</p>
						<p><span style="font-weight:bold">内容过滤策略：更灵活的安全措施</span><br/>
						　　如果说IP屏蔽策略适用于完全屏蔽、长期生效的固定IP，那么内容过滤策略将会是更灵活的解决方案。该策略允许使用正则表达式来匹配规律性的字符串，并且版主可以根据情况选择替换字符串或者直接拒绝留言。</p>
						<p><span style="font-weight:bold">其它功能</span><br/>
						　　留言审核、悄悄话、邮件通知等。</p>
						<br/><br/>
					</td>
				</tr>
			</table>		
		
		</td>
	</tr>
</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
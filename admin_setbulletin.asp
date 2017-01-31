<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_setbulletin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 发布置顶公告</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function sfocus()
	{
		if (!form6.abulletin.modified)
		{
			form6.abulletin.focus();
			form6.abulletin.select();
		}
	}
	</script>
</head>

<body<%=bodylimit%> onload="sfocus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("管理")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	Call CreateConn(cn)

	rs.Open Replace(sql_adminsetbulletin,"{0}",adminid),cn,,,1
	dim tbul,tflag
	tflag=rs("declareflag")
	tbul=rs("declare")

	if isnull(tbul) then
		tbul=""
	else
		tbul=HtmlEncode(tbul)
	end if
	if tbul="" then tflag=(web_adminlimit and adminlimit)

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	%>

	<div class="region">
		<h3 class="title">发布置顶公告</h3>
		<div class="content">
			<form method="post" action="admin_savebulletin.asp" name="form6" onsubmit="form6.submit1.disabled=true;">
				<input type="hidden" name="user" value="<%=ruser%>" />
				公告内容：<br/><textarea name="abulletin" id="abulletin" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" rows="<%=ReplyTextHeight%>"><%=tbul%></textarea>
				<!-- #include file="include/template/ubbtoolbar.inc" -->
				<%if web_AdminUBBSupport then ShowUbbToolBar(2)%>
				<p>
				<%if web_AdminHTMLSupport then%><input type="checkbox" name="html2" id="html2" value="1"<%=cked(CBool(tflag AND 1))%> /><label for="html2">支持HTML标记</label><br/><%end if%>
				<%if web_AdminUBBSupport then%><input type="checkbox" name="ubb2" id="ubb2" value="1"<%=cked(CBool(tflag AND 2))%> /><label for="ubb2">支持UBB标记</label><br/><%end if%>
				<%if web_AdminAllowNewLine then%><input type="checkbox" name="newline2" id="newline2" value="1"<%=cked(CBool(tflag AND 4))%> /><label for="newline2">不支持HTML和UBB标记时允许回车换行</label><%end if%>
				</p>
				<div class="command"><input value="更新数据" type="submit" name="submit1" id="submit1" /></div>
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
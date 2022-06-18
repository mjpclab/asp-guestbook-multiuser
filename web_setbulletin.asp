<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/web.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_setbulletin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%Response.Expires=-1%>

<!-- #include file="include/template/dtd.inc" -->
<html lang="zh-CN">
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster管理中心 发布系统公告</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->

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

<body onload="sfocus();">

<div id="outerborder" class="outerborder">

	<%Call WebInitHeaderData("","Webmaster管理中心","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	Call CreateConn(cn)

	rs.Open sql_websetbulletin,cn,,,1
	dim tbul,tflag

	tflag=rs("bulletinflag")
	tbul="" & rs("webbulletin") & ""

	if isnull(tbul) then
		tbul=""
	else
		tbul=HtmlEncode(tbul)
	end if

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	%>

	<div class="region">
		<h3 class="title">发布系统公告</h3>
		<div class="content">
			<form method="post" action="web_savebulletin.asp" name="form6" onsubmit="form6.submit1.disabled=true;">
			公告内容：<br/><textarea name="abulletin" id="abulletin" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" rows="<%=ReplyTextHeight%>"><%=tbul%></textarea>
			<!-- #include file="include/template/ubbtoolbar.inc" -->
			<%ShowUbbToolBar(1)%>
			<p>
			<input type="checkbox" name="html2" id="html2" value="1"<%=cked(CBool(tflag AND 1))%> /><label for="html2">支持HTML标记</label><br/>
			<input type="checkbox" name="ubb2" id="ubb2" value="1"<%=cked(CBool(tflag AND 2))%> /><label for="ubb2">支持UBB标记</label><br/>
			<input type="checkbox" name="newline2" id="newline2" value="1"<%=cked(CBool(tflag AND 4))%> /><label for="newline2">不支持HTML和UBB标记时允许回车换行</label><br/><br/>

			<input type="checkbox" name="pub_at_face" id="pub_at_face" value="1"<%=cked(CBool(tflag AND 16))%> /><label for="pub_at_face">在首页发布</label><br/>
			<input type="checkbox" name="pub_at_function" id="pub_at_function" value="1"<%=cked(CBool(tflag AND 32))%> /><label for="pub_at_function">在功能页发布(注册、维护……)</label><br/>
			<input type="checkbox" name="pub_at_index" id="pub_at_index" value="1"<%=cked(CBool(tflag AND 64))%> /><label for="pub_at_index">在用户留言、搜索页发布</label><br/>
			<input type="checkbox" name="pub_at_search" id="pub_at_search" value="1"<%=cked(CBool(tflag AND 128))%> /><label for="pub_at_search">在用户管理员首页发布</label>
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
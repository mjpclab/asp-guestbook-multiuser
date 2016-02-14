<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%Response.Expires=-1%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster管理中心 高级删除</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function del_all_user_warning(frm)
	{
		if(confirm('警告！删除的用户和数据将不能恢复！\n确实要执行删除操作吗？'))
			if (confirm('请再次确认是否要删除用户？'))
				{frm.submit1.disabled=true;return true;}
			else
				return false;
		else
			return false;
	}
	</script>
</head>

<body>

<div id="outerborder" class="outerborder">

	<%Call WebInitHeaderData("","Webmaster管理中心","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<div class="region form-region">
		<h3 class="title">高级删除</h3>
		<div class="content">
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除指定日期和时间前的留言，包括此日期/时间。<br/>
				<input type="hidden" name="option" value="1" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />(时间默认为0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除指定日期和时间后的留言，包括此日期/时间。<br/>
				<input type="hidden" name="option" value="2" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />(时间默认为0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除最靠前(最老)的n条留言，请输入n的值：<br/>
				<input type="hidden" name="option" value="3" />
				<input type="text" name="iparam" maxlength="10" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除最靠后(最新)的n条留言，请输入n的值：<br/>
				<input type="hidden" name="option" value="4" />
				<input type="text" name="iparam" maxlength="10" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除称呼中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="5" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除标题中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="6" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除访客留言中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="7" />
				<input type="text" name="iparam" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除版主回复中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="8" />
				<input type="text" name="iparam" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除版主回复中包含特定字符串的回复：<br/>
				<input type="hidden" name="option" value="9" />
				<input type="text" name="iparam" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除置顶公告中包含特定字符串的公告：<br/>
				<input type="hidden" name="option" value="12" />
				<input type="text" name="iparam" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除全部留言，清空留言本，请谨慎使用：
				<input type="hidden" name="option" value="10" />
				<input type="hidden" name="iparam" value="DEL_ALL" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除全部版主回复，请谨慎使用：
				<input type="hidden" name="option" value="11" />
				<input type="hidden" name="iparam" value="DEL_ALL_REPLY" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除全部置顶公告，请谨慎使用：
				<input type="hidden" name="option" value="13" />
				<input type="hidden" name="iparam" value="DEL_ALL_DECLARE" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>

			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)">
				删除注册日期在指定日期和时间前的用户，包括此日期/时间：<br/>
				<input type="hidden" name="option" value="14" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />(时间默认为0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)" id=form1 name=form1>
				删除最后登录日期在指定日期和时间前的用户，包括此日期/时间：<br/>
				<input type="hidden" name="option" value="15" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />(时间默认为0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)" id=form2 name=form2>
				删除最后留言日期在指定日期和时间前的用户，包括此日期/时间：<br/>
				<input type="hidden" name="option" value="16" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />(时间默认为0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)">
				删除从未登录过的用户，请谨慎使用：
				<input type="hidden" name="option" value="17" />
				<input type="hidden" name="iparam" value="DEL_NeverLogined_USER" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)">
				删除全部注册用户，请谨慎使用：
				<input type="hidden" name="option" value="18" />
				<input type="hidden" name="iparam" value="DEL_ALL_USER" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
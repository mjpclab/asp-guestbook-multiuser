<!-- #include file="webconfig.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusFindkey=false then
	Response.Redirect "web_err.asp?number=2"
	Response.End
end if
%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=web_BookName%> 找回密码步骤1</title>
	<!-- #include file="inc_stylesheet.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.user.value.length===0) {alert('请输入用户名。'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)===false) {alert('用户名中只能包含英文字母、数字和下划线。'); frm.user.select(); return false;}
		else {frm.submit1.disabled=true; return true;}
	}
	</script>
</head>

<body onload="findform.user.select();">

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">找回密码</a> &gt;&gt; 步骤1</div>
</div>

<%
set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="include/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/web_guest_func.inc" -->

<div class="region form-region">
	<h3 class="title">找回密码 步骤1</h3>
	<div class="content">
	<form name="findform" method="post" action="findkey2.asp" onsubmit="return submitcheck(this)">
		<div class="field">
			<span class="label">用户名：</span>
			<span class="value"><input type="text" name="user" maxlength="32" size="32" title="只能包含英文、数字和下划线，最大32位" /></span>
		</div>
		<div class="command"><input type="submit" name="submit1" value="下一步" />　　<input type="reset" value="重新填写" /></div>
	</form>
	</div>
</div>

</div>

<!-- #include file="include/footer.inc" -->
</body>
</html>
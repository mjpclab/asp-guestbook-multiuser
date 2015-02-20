<!-- #include file="webconfig.asp" -->
<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusFindkey=false then
	Response.Redirect "web_err.asp?number=2"
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> 找回密码步骤1</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.user.value=='') {alert('请输入用户名。'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)==false) {alert('用户名中只能包含英文字母、数字和下划线。'); frm.user.select(); return false;}
		else {frm.submit1.disabled=true; return true;}
	}
	</script>
</head>

<body onload="findform.user.select();">

<div id="outerborder" class="outerborder">

<table cellpadding="2" cellspacing="0" style="width:100%; border-width:0px; border-collapse:collapse;">
	<tr style="height:60px;">
		<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);" nowrap="nowrap">
			<%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">找回密码</a> &gt;&gt; 步骤1
		</td>
	</tr>		
	<tr>
		<td>
			<%			
			set cn=server.CreateObject("ADODB.Connection")
			CreateConn cn,dbtype
			
			dim sys_bul_flag
			sys_bul_flag=32
			%>
			<!-- #include file="sysbulletin.inc" -->
			<%cn.Close : set cn=nothing%>
		</td>
	</tr>
	<tr>
		<td style="width:100%">
		<!-- #include file="func_web.inc" -->
		</td>
	</tr>
		<td style="width:100%; text-align:center; padding:20px 0px; color:<%=TableContentColor%>;">
			<form name="findform" method="post" action="findkey2.asp" onsubmit="return submitcheck(this)">
				<br/>
				用户名：　　<input type="text" name="user" maxlength="32" size="32" title="只能包含英文、数字和下划线，最大32位" /><br/><br/>
				<br/><input type="submit" name="submit1" value="下一步" />　　<input type="reset" value="重新填写" />
			</form>
		</td>
	</tr>
</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
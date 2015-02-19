<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusUnreg=false then
	Response.Redirect "web_err.asp?number=5"
	Response.End
end if

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=web_BookName%> 自删帐号</title>

	<script type="text/javascript">
	//<![CDATA[
	
	function submitcheck(frm)
	{
		if (frm.user.value=='') {alert('请输入用户名。'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)==false) {alert('用户名中只能包含英文字母、数字和下划线。'); frm.user.focus(); return false;}
		else if (frm.pass1.value=='') {alert('请输入密码。'); frm.pass1.focus(); return false;}
		else if (frm.vcode && frm.vcode.value=='') {alert('请输入验证码。'); frm.vcode.focus(); return false;}
		else return confirm('删除用户名将同时清除全部留言、回复及公告，确实要继续吗？');
	}
	
	//]]>
	</script>
	
	<!-- #include file="style.asp" -->
</head>

<body onload="unregform.user.focus();">

<div id="outerborder" class="outerborder">

<table cellpadding="2" cellspacing="0" style="width:100%; border-width:0px; border-collapse:collapse;">
	<tr style="height:60px;">
		<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);" nowrap="nowrap">
			<%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; 自删帐号
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
			<form name="unregform" method="post" action="submitunreg.asp" onsubmit="return submitcheck(this)">

			<table cellpadding="5" cellspacing="0" style="margin-left:auto; margin-right:auto; border-width:0px;">
			<tr>
				<td>用户名：</td>
				<td><input type="text" name="user" maxlength="32" size="32" title="只能包含英文、数字和下划线，最大32位" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>密码：</td>
				<td><input type="password" name="pass1" maxlength="32" size="32" /></td>
			</tr>
			<%if VcodeCount>0 then%>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>验证码：</td>
				<td><input type="text" name="vcode" size="10" /> <img src="web_show_vcode.asp" class="vcode" onclick="this.src=this.src" /></td>
			</tr>
			<%end if%>
			<tr style="height:20px;"><td></td></tr>
			<tr>
				<td colspan="2" style="text-align:center;">
					<input type="submit" value="提交" />　　<input type="reset" value="重写" />
				</td>
			</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
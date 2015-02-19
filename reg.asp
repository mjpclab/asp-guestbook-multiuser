<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusReg=false then
	Response.Redirect "web_err.asp?number=1"
	Response.End
end if

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=web_BookName%> 创建留言本</title>

	<script type="text/javascript">
	//<![CDATA[
	
	function submitcheck(frm)
	{
		if (frm.user.value=='') {alert('请输入用户名。'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)==false) {alert('用户名中只能包含英文字母、数字和下划线。'); frm.user.select(); return false;}
		else if (frm.pass1.value=='') {alert('请输入密码。'); frm.pass1.focus(); return false;}
		else if (frm.pass2.value=='') {alert('请输入确认密码。'); frm.pass2.focus(); return false;}
		else if (frm.pass1.value!=frm.pass2.value) {alert('密码不一致，请检查。'); frm.pass2.select(); return false;}
		else if (frm.question.value=='') {alert('请输入找回密码问题。'); frm.question.focus(); return false;}
		else if (frm.key.value=='') {alert('请输入找回密码答案。'); frm.key.focus(); return false;}
		else if (frm.vcode && frm.vcode.value=='') {alert('请输入验证码。'); frm.vcode.focus(); return false;}
		else return true;
	}
	
	<!-- #include file="xmlhttp.inc" -->
	function checkuser(frm,showChecking)
	{
		/*
		if (frm.user.value=='') {alert('请输入用户名。'); frm.user.focus(); return;}
		else if (/^\w+$/.test(frm.user.value)==false) {alert('用户名中只能包含英文字母、数字和下划线。'); frm.user.select(); return;}
		else window.open('checkuser.asp?user=' + frm.user.value,'win_checkuser','width=300,height=185,menubar=no,toolbar=no,status=no,scrollbars=no,resizable=no');
		*/
		var username=frm.user.value;
		var divCheckuser=document.getElementById('divCheckuser');
		if(divCheckuser)
		{
			while(divCheckuser.childNodes.length>0) divCheckuser.removeChild(divCheckuser.lastChild);
			if(username=='') setPureText(divCheckuser,'请输入用户名。');
			else if(!/^\w+$/.test(username)) setPureText(divCheckuser,'用户名只能包含英文字母、数字和下划线。');
			else
			{
				window.xmlHttp=createXmlHttp();
				if(xmlHttp)
				{
					if(showChecking)
						{setPureText(divCheckuser,'正在检测用户名，请稍候……');}
					else
						{clearChildren(divCheckuser);}
						
					xmlHttp.abort();
					xmlHttp.onreadystatechange=checkuserComplete;
					xmlHttp.open('GET','checkuser.asp?user=' + username);
					xmlHttp.send(null);
				}
			}
		}
	}
	
	function checkuserComplete()
	{
		if(xmlHttp.readyState==4 && xmlHttp.status==200)
		{
			var divCheckuser=document.getElementById('divCheckuser');
			if(divCheckuser) setPureText(divCheckuser,xmlHttp.responseText);
			xmlHttp.abort();
		}
	}
	
	//]]>
	</script>
	
	<!-- #include file="style.asp" -->
</head>

<body onload="if(regform.user.value=='')regform.user.focus();">

<div id="outerborder" class="outerborder">

<table cellpadding="2" cellspacing="0" style="width:100%; border-width:0px; border-collapse:collapse;">
	<tr style="height:60px;">
		<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);" nowrap="nowrap">
			<%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; 创建留言本
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
		<td style="width:100%; text-align:center; color:<%=TableContentColor%>;">
			<form name="regform" method="post" action="submitreg.asp" onsubmit="return submitcheck(this)">

			<table cellpadding="5" cellspacing="0" style="margin:20px auto; border-style:none; border-width:0px;">
				<td>用户名*</td>
				<td>
					<input type="text" name="user" maxlength="32" size="18" title="只能包含英文字母、数字和下划线，最大32位" onchange="checkuser(this.form,false)" />
					<input type="button" value="检测帐号" onclick="checkuser(this.form,true)" />
				</td>
			</tr>
			<tr>
				<td></td>
				<td><div id="divCheckuser" style="width:230px; height:10ex;"></div></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>密码*</td>
				<td><input type="password" name="pass1" maxlength="32" size="32" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>确认密码*</td>
				<td><input type="password" name="pass2" maxlength="32" size="32" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>版主昵称</td>
				<td><input type="text" name="nick" maxlength="64" size="32" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>提示问题*</td>
				<td><input type="text" name="question" maxlength="32" size="32" title="用于找回密码" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>问题答案*</td>
				<td><input type="text" name="key" maxlength="32" size="32" /></td>
			</tr>
			<%if VcodeCount>0 then%>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>验证码</td>
				<td><input type="text" name="vcode" size="10" /> <img src="web_show_vcode.asp" class="vcode" onclick="this.src=this.src" /></td>
			</tr>
			<%end if%>
			<tr style="height:20px;"><td></td></tr>
			<tr>
				<td colspan="2" style="text-align:center;">
					<input type="submit" value="提交" style="width:80px; height:25px;" />　　<input type="reset" value="清空" style="width:80px; height:25px;" />
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
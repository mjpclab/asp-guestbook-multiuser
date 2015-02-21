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
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> 创建留言本</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.user.value.length===0) {alert('请输入用户名。'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)==false) {alert('用户名中只能包含英文字母、数字和下划线。'); frm.user.select(); return false;}
		else if (frm.pass1.value.length===0) {alert('请输入密码。'); frm.pass1.focus(); return false;}
		else if (frm.pass2.value.length===0) {alert('请输入确认密码。'); frm.pass2.focus(); return false;}
		else if (frm.pass1.value!=frm.pass2.value) {alert('密码不一致，请检查。'); frm.pass2.select(); return false;}
		else if (frm.question.value.length===0) {alert('请输入找回密码问题。'); frm.question.focus(); return false;}
		else if (frm.key.value.length===0) {alert('请输入找回密码答案。'); frm.key.focus(); return false;}
		else if (frm.vcode && frm.vcode.value.length===0) {alert('请输入验证码。'); frm.vcode.focus(); return false;}
		else return true;
	}
	
	<!-- #include file="xmlhttp.inc" -->
	function checkuser(frm,showChecking)
	{
		/*
		if (frm.user.value.length===0) {alert('请输入用户名。'); frm.user.focus(); return;}
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
	</script>
</head>

<body onload="if(regform.user.value.length===0)regform.user.focus();">

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; 创建留言本</div>
</div>

<%
set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="func_web.inc" -->

<div class="region form-region">
	<h3 class="title">创建留言本</h3>
	<div class="content">
	<form name="regform" method="post" action="submitreg.asp" onsubmit="return submitcheck(this)">
		<div class="field">
			<span class="label">用户名<span class="required">*</span></span>
			<span class="value">
				<input type="text" name="user" maxlength="32" size="18" title="只能包含英文字母、数字和下划线，最大32位" onchange="checkuser(this.form,false)" />
				<input type="button" value="检测帐号" onclick="checkuser(this.form,true)" />
			</span>
		</div>
		<div class="field">
			<span class="label">&nbsp;</span>
			<span class="value"><div id="divCheckuser">&nbsp;</div></span>
		</div>
		<div class="field">
			<span class="label">密码<span class="required">*</span></span>
			<span class="value"><input type="password" name="pass1" maxlength="32" size="32" /></span>
		</div>
		<div class="field">
			<span class="label">确认密码<span class="required">*</span></span>
			<span class="value"><input type="password" name="pass2" maxlength="32" size="32" /></span>
		</div>
		<div class="field">
			<span class="label">版主昵称</span>
			<span class="value"><input type="text" name="nick" maxlength="64" size="32" /></span>
		</div>
		<div class="field">
			<span class="label">提示问题<span class="required">*</span></span>
			<span class="value"><input type="text" name="question" maxlength="32" size="32" title="用于找回密码" /></span>
		</div>
		<div class="field">
			<span class="label">问题答案<span class="required">*</span></span>
			<span class="value"><input type="text" name="key" maxlength="32" size="32" /></span>
		</div>
		<%if VcodeCount>0 then%>
		<div class="field">
			<span class="label">验证码<span class="required">*</span></span>
			<span class="value"><input type="text" name="vcode" size="10" /><img class="captcha" src="web_show_vcode.asp"/></span>
		</div>
		<%end if%>
		<div class="command"><input type="submit" value="提交" style="width:80px; height:25px;" />　　<input type="reset" value="清空" style="width:80px; height:25px;" /></div>
	</form>
	</div>
</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>
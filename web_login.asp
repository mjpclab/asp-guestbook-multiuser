<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� �������ĵ�¼</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
	<link rel="stylesheet" type="text/css" href="web_adminstyle.css"/>
	<!-- #include file="style.asp" -->
	<!-- #include file="adminstyle.asp" -->
	<!-- #include file="web_adminstyle.asp" -->
	
	<script type="text/javascript">
	function submitCheck(obj)
	{
		if(obj.iadminpass.value.length===0)
		{
			alert('���������롣');
			obj.iadminpass.focus();
			return false;
		}
		else if(obj.ivcode && obj.ivcode.value.length===0)
		{
			alert('��������֤�롣');
			obj.ivcode.focus();
			return false;
		}
		else
		{
			obj.submit1.disabled=true;
			return true;
		}
	}
	</script>
</head>

<body onload="if(form5.iadminpass.value=='')form5.iadminpass.focus()">

<div class="region form-region narrow-region">
	<h3 class="title">Webmaster �������ĵ�¼</h3>
	<div class="content">
		<form method="post" action="web_login_verify.asp" name="form5" onsubmit="return submitCheck(this);">
			<div class="field">
				<span class="label">���룺</span>
				<span class="value"><input type="password" name="iadminpass" maxlength="32" autofocus="autofocus"/></span>
			</div>
			<%if VcodeCount>0 then%>
			<div class="field">
				<span class="label">��֤�룺</span>
				<span class="value"><input type="text" name="ivcode" autocomplete="off" /><img class="captcha" src="web_show_vcode.asp"/></span>
			</div>
			<%end if%>
			<div class="command">
				<input type="submit" value="��¼" name="submit1" />
			</div>
			</form>
	</div>
</div>

</body>
</html>